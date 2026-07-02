// ============================================================
// WEDDING GUEST TRACKER — Google Apps Script Backend
// ============================================================
// Deploy as Web App:
//   Extensions > Apps Script > Deploy > New Deployment
//   Type: Web App | Execute as: Me | Who has access: Anyone
// ============================================================

const SHEET_NAME_GUESTS  = "guests";
const SHEET_NAME_DRIVERS = "drivers";

const GUEST_HEADERS = [
  "group_id","group_name","individual_names","pax",
  "phone_primary","phone_backup","pickup_location",
  "journey_origin","journey_destination","arrival_date","arrival_time",
  "transport_type","transport_name","transport_number","pnr",
  "car_type","car_no","driver_name","driver_phone",
  "notes","status","journey_type","dispatched",
  "driver_arrived","guest_in_car","dropped_off","car_returned",
  "last_refreshed","live_status_text","deleted","ticket_url"
];

const DRIVER_HEADERS = [
  "driver_id","driver_name","driver_phone",
  "car_type","car_no","capacity","assigned"
];

// ── Formula Injection Protection ─────────────────────────────
// Prefix dangerous characters with apostrophe to prevent CSV injection
function safeCellValue(val) {
  if (val === "" || val === null || val === undefined) return "";
  var s = String(val);
  // Characters that trigger formula execution in Sheets
  if (/^[=+\-@]/.test(s)) return "'" + s;
  return s;
}
function sanitizeRow(row) {
  return row.map(function(v) { return safeCellValue(v); });
}

function cors(output) {
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === "getGuests")  return cors(getGuests());
    if (action === "getDrivers") return cors(getDrivers());
    return cors({ error: "Unknown action" });
  } catch(err) { return cors({ error: err.message }); }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;
    if (action === "addGuest")          return cors(addGuest(payload.data));
    if (action === "updateGuest")       return cors(updateGuest(payload.data));
    if (action === "updateGuestStatus") return cors(updateGuestStatus(payload.group_id, payload.field, payload.value));
    if (action === "updateLiveStatus")  return cors(updateLiveStatus(payload.group_id, payload.live_status_text));
    if (action === "deleteGuest")       return cors(deleteGuest(payload.group_id));
    if (action === "saveTicket")        return cors(saveTicket(payload.group_id, payload.data, payload.mime, payload.name));
    if (action === "initSheet")         return cors(initSheet());
    return cors({ error: "Unknown action" });
  } catch(err) { return cors({ error: err.message }); }
}

function initSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();

  var gSheet = ss.getSheetByName(SHEET_NAME_GUESTS);
  if (!gSheet) gSheet = ss.insertSheet(SHEET_NAME_GUESTS);
  if (gSheet.getLastRow() === 0) {
    gSheet.appendRow(GUEST_HEADERS);
    gSheet.getRange(1,1,1,GUEST_HEADERS.length).setFontWeight("bold");
  }
  // Always force plain text on date + time columns — runs every app load, safe to repeat
  var dateCol = GUEST_HEADERS.indexOf("arrival_date") + 1;
  var timeCol = GUEST_HEADERS.indexOf("arrival_time") + 1;
  var lastRow = Math.max(gSheet.getLastRow(), 2);
  gSheet.getRange(2, dateCol, lastRow, 1).setNumberFormat("@STRING@");
  gSheet.getRange(2, timeCol, lastRow, 1).setNumberFormat("@STRING@");

  // Add ticket_url column if missing (safe to run on existing sheets)
  var headers = gSheet.getRange(1, 1, 1, gSheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("ticket_url") === -1) {
    var nextCol = gSheet.getLastColumn() + 1;
    gSheet.getRange(1, nextCol).setValue("ticket_url").setFontWeight("bold");
  }

  var dSheet = ss.getSheetByName(SHEET_NAME_DRIVERS);
  if (!dSheet) dSheet = ss.insertSheet(SHEET_NAME_DRIVERS);
  if (dSheet.getLastRow() === 0) {
    dSheet.appendRow(DRIVER_HEADERS);
    dSheet.getRange(1,1,1,DRIVER_HEADERS.length).setFontWeight("bold");
  }
  return { success: true, message: "Sheet initialised", tz: tz };
}

// ── Format any cell value to a clean string ──────────────────
function formatCell(value, header, tz) {
  if (value === "" || value === null || value === undefined) return "";

  if (value instanceof Date) {
    // 1899 epoch = empty cell in Sheets
    if (value.getFullYear() < 1900) return "";
    // Always use Utilities.formatDate with the spreadsheet timezone
    // This is the only correct way — it reads the Date in the sheet's local timezone
    if (header === "arrival_time")   return Utilities.formatDate(value, tz, "HH:mm");
    if (header === "arrival_date")   return Utilities.formatDate(value, tz, "yyyy-MM-dd");
    if (header === "last_refreshed") return Utilities.formatDate(value, tz, "yyyy-MM-dd HH:mm");
    return Utilities.formatDate(value, tz, "yyyy-MM-dd HH:mm");
  }

  if (typeof value === "string") {
    // Already a plain string — validate format
    if (header === "arrival_time") {
      // Accept HH:MM or HH:MM:SS, strip seconds
      var tm = value.match(/^(\d{1,2}:\d{2})/);
      if (tm) return tm[1].length === 4 ? "0" + tm[1] : tm[1]; // pad single digit hour
      return "";
    }
    if (header === "arrival_date") {
      // Accept YYYY-MM-DD only
      var dm = value.match(/^(\d{4}-\d{2}-\d{2})/);
      if (dm) { if (parseInt(dm[1].slice(0,4)) < 1900) return ""; return dm[1]; }
      return "";
    }
    // ISO datetime string that slipped through
    if (/^\d{4}-\d{2}-\d{2}T/.test(value)) {
      var dt = new Date(value);
      if (isNaN(dt.getTime()) || dt.getFullYear() < 1900) return "";
      if (header === "arrival_time") return Utilities.formatDate(dt, tz, "HH:mm");
      if (header === "arrival_date") return Utilities.formatDate(dt, tz, "yyyy-MM-dd");
      return value;
    }
    return value;
  }

  return String(value);
}

function getGuests() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var tz    = ss.getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(SHEET_NAME_GUESTS);
  if (!sheet || sheet.getLastRow() <= 1) return { guests: [] };
  // Read up to max(GUEST_HEADERS.length, lastColumn) to handle sheets with extra columns.
  // We only map the columns we know about via GUEST_HEADERS — extras are ignored safely.
  var numCols = Math.max(GUEST_HEADERS.length, sheet.getLastColumn());
  var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, numCols).getValues();
  var guests = rows
    .filter(function(r){ return r[0] !== ""; })
    .map(function(r){
      var obj = {};
      GUEST_HEADERS.forEach(function(h, i){ obj[h] = formatCell(r[i], h, tz); });
      return obj;
    });
  return { guests: guests };
}

function getDrivers() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var tz    = ss.getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(SHEET_NAME_DRIVERS);
  if (!sheet || sheet.getLastRow() <= 1) return { drivers: [] };
  var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, DRIVER_HEADERS.length).getValues();
  var drivers = rows
    .filter(function(r){ return r[0] !== ""; })
    .map(function(r){
      var obj = {};
      DRIVER_HEADERS.forEach(function(h, i){ obj[h] = formatCell(r[i], h, tz); });
      return obj;
    });
  return { drivers: drivers };
}

function addGuest(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
  if (!sheet) return { error: "Guests sheet not found. Run initSheet first." };
  data.group_id       = "G" + Date.now();
  data.status         = "active";
  data.journey_type   = data.journey_type || "arrival";
  data.dispatched     = "no";
  data.driver_arrived = "no";
  data.guest_in_car   = "no";
  data.dropped_off    = "no";
  data.car_returned   = "no";
  data.last_refreshed     = "";
  data.live_status_text   = "";
  data.deleted            = "no";
  data.ticket_url         = data.ticket_url || "";
  var row = sanitizeRow(GUEST_HEADERS.map(function(h){ return data[h] || ""; }));
  sheet.appendRow(row);
  return { success: true, group_id: data.group_id };
}

function updateGuest(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
  var rowIndex = findGuestRow(sheet, data.group_id);
  if (!rowIndex) return { error: "Guest not found" };
  var row = sanitizeRow(GUEST_HEADERS.map(function(h){ return data[h] !== undefined ? data[h] : ""; }));
  sheet.getRange(rowIndex, 1, 1, GUEST_HEADERS.length).setValues([row]);
  return { success: true };
}

function updateGuestStatus(group_id, field, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
  var rowIndex = findGuestRow(sheet, group_id);
  if (!rowIndex) return { error: "Guest not found" };
  var colIndex = GUEST_HEADERS.indexOf(field) + 1;
  if (colIndex === 0) return { error: "Field not found: " + field };
  sheet.getRange(rowIndex, colIndex).setValue(value);
  return { success: true };
}

function updateLiveStatus(group_id, live_status_text) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var tz    = ss.getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(SHEET_NAME_GUESTS);
  var rowIndex = findGuestRow(sheet, group_id);
  if (!rowIndex) return { error: "Guest not found" };
  var tsCol   = GUEST_HEADERS.indexOf("last_refreshed") + 1;
  var textCol = GUEST_HEADERS.indexOf("live_status_text") + 1;
  var ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
  sheet.getRange(rowIndex, tsCol).setValue(ts);
  sheet.getRange(rowIndex, textCol).setValue(live_status_text);
  return { success: true };
}

function deleteGuest(group_id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
  var rowIndex = findGuestRow(sheet, group_id);
  if (!rowIndex) return { error: "Guest not found" };
  // Trash any attached Drive ticket file
  var ticketCol = GUEST_HEADERS.indexOf("ticket_url") + 1;
  var ticketUrl = sheet.getRange(rowIndex, ticketCol).getValue();
  if (ticketUrl) {
    try {
      var m = ticketUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (m) DriveApp.getFileById(m[1]).setTrashed(true);
    } catch(e) { /* already gone or no access */ }
  }
  var colIndex = GUEST_HEADERS.indexOf("deleted") + 1;
  sheet.getRange(rowIndex, colIndex).setValue("yes");
  return { success: true };
}

function findGuestRow(sheet, group_id) {
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === group_id) return i + 2;
  }
  return null;
}

// ── Drive Ticket Storage ──────────────────────────────────────
// Saves ticket file to a shared Drive folder and stores the URL
// in the ticket_url column of the guests sheet.

function getTicketFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty("TICKET_FOLDER_ID");
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch(e) { /* folder deleted, recreate */ }
  }
  // Create folder in root of My Drive
  var folder = DriveApp.createFolder("Wedding Ticket Attachments");
  // Make folder viewable by anyone with link
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty("TICKET_FOLDER_ID", folder.getId());
  return folder;
}

function saveTicket(group_id, base64Data, mimeType, fileName) {
  if (!group_id || !base64Data || !mimeType) {
    return { error: "Missing required fields: group_id, data, mime" };
  }

  try {
    var folder = getTicketFolder();

    // Decode base64 to blob
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType, fileName || ("ticket_" + group_id));

    // Delete any existing ticket for this guest (keep Drive clean)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
    var rowIndex = findGuestRow(sheet, group_id);
    var ticketCol = GUEST_HEADERS.indexOf("ticket_url") + 1; // declared once, used twice below
    if (rowIndex) {
      var existingUrl = sheet.getRange(rowIndex, ticketCol).getValue();
      if (existingUrl) {
        try {
          var match = existingUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
          if (match) DriveApp.getFileById(match[1]).setTrashed(true);
        } catch(e) { /* old file may already be gone, ignore */ }
      }
    }

    // Save new file to Drive folder
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileId = file.getId();
    var viewUrl = "https://drive.google.com/file/d/" + fileId + "/preview";

    // Persist ticket_url in the Sheet
    if (rowIndex) {
      sheet.getRange(rowIndex, ticketCol).setValue(viewUrl);
    }

    return { success: true, ticket_url: viewUrl, file_id: fileId };

  } catch(err) {
    return { error: "Drive save failed: " + err.message };
  }
}