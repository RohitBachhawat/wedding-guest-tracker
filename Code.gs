// ============================================================
// WEDDING GUEST TRACKER — Google Apps Script Backend
// ============================================================
// Deploy this as a Web App:
//   Extensions > Apps Script > Deploy > New Deployment
//   Type: Web App | Execute as: Me | Who has access: Anyone
// ============================================================

const SHEET_NAME_GUESTS  = "guests";
const SHEET_NAME_DRIVERS = "drivers";

const GUEST_HEADERS = [
  "group_id","group_name","individual_names","pax",
  "phone_primary","phone_backup","pickup_location",
  "journey_origin","arrival_date","arrival_time",
  "transport_type","transport_name","transport_number","pnr",
  "car_type","car_token","driver_name","driver_phone",
  "notes","status","dispatched","last_refreshed","live_status_text"
];

const DRIVER_HEADERS = [
  "driver_id","driver_name","driver_phone",
  "car_type","car_token","capacity","assigned"
];

// ── CORS helper ──────────────────────────────────────────────
function cors(output) {
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Router ───────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === "getGuests")  return cors(getGuests());
    if (action === "getDrivers") return cors(getDrivers());
    return cors({ error: "Unknown action" });
  } catch(err) {
    return cors({ error: err.message });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    if (action === "addGuest")          return cors(addGuest(payload.data));
    if (action === "updateGuest")       return cors(updateGuest(payload.data));
    if (action === "updateGuestStatus") return cors(updateGuestStatus(payload.group_id, payload.field, payload.value));
    if (action === "updateLiveStatus")  return cors(updateLiveStatus(payload.group_id, payload.live_status_text));
    if (action === "initSheet")         return cors(initSheet());
    return cors({ error: "Unknown action" });
  } catch(err) {
    return cors({ error: err.message });
  }
}

// ── Init Sheet (creates headers if sheet is empty) ───────────
function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let gSheet = ss.getSheetByName(SHEET_NAME_GUESTS);
  if (!gSheet) gSheet = ss.insertSheet(SHEET_NAME_GUESTS);
  if (gSheet.getLastRow() === 0) {
    gSheet.appendRow(GUEST_HEADERS);
    gSheet.getRange(1, 1, 1, GUEST_HEADERS.length).setFontWeight("bold");
  }

  let dSheet = ss.getSheetByName(SHEET_NAME_DRIVERS);
  if (!dSheet) dSheet = ss.insertSheet(SHEET_NAME_DRIVERS);
  if (dSheet.getLastRow() === 0) {
    dSheet.appendRow(DRIVER_HEADERS);
    dSheet.getRange(1, 1, 1, DRIVER_HEADERS.length).setFontWeight("bold");
  }

  return { success: true, message: "Sheet initialised" };
}

// ── Read guests ───────────────────────────────────────────────
function getGuests() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
  if (!sheet || sheet.getLastRow() <= 1) return { guests: [] };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, GUEST_HEADERS.length).getValues();
  const guests = rows
    .filter(r => r[0] !== "")
    .map(r => {
      const obj = {};
      GUEST_HEADERS.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
  return { guests };
}

// ── Read drivers ──────────────────────────────────────────────
function getDrivers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_DRIVERS);
  if (!sheet || sheet.getLastRow() <= 1) return { drivers: [] };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, DRIVER_HEADERS.length).getValues();
  const drivers = rows
    .filter(r => r[0] !== "")
    .map(r => {
      const obj = {};
      DRIVER_HEADERS.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
  return { drivers };
}

// ── Add guest ─────────────────────────────────────────────────
function addGuest(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
  if (!sheet) return { error: "Guests sheet not found. Run initSheet first." };
  data.group_id       = "G" + Date.now();
  data.status         = "active";
  data.dispatched     = "no";
  data.last_refreshed = "";
  data.live_status_text = "";
  const row = GUEST_HEADERS.map(h => data[h] || "");
  sheet.appendRow(row);
  return { success: true, group_id: data.group_id };
}

// ── Update full guest row ─────────────────────────────────────
function updateGuest(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
  const rowIndex = findGuestRow(sheet, data.group_id);
  if (!rowIndex) return { error: "Guest not found" };
  const row = GUEST_HEADERS.map(h => data[h] !== undefined ? data[h] : "");
  sheet.getRange(rowIndex, 1, 1, GUEST_HEADERS.length).setValues([row]);
  return { success: true };
}

// ── Update single field (status, dispatched) ──────────────────
function updateGuestStatus(group_id, field, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
  const rowIndex = findGuestRow(sheet, group_id);
  if (!rowIndex) return { error: "Guest not found" };
  const colIndex = GUEST_HEADERS.indexOf(field) + 1;
  if (colIndex === 0) return { error: "Field not found" };
  sheet.getRange(rowIndex, colIndex).setValue(value);
  return { success: true };
}

// ── Update live status after Gemini refresh ───────────────────
function updateLiveStatus(group_id, live_status_text) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_GUESTS);
  const rowIndex = findGuestRow(sheet, group_id);
  if (!rowIndex) return { error: "Guest not found" };

  const tsCol   = GUEST_HEADERS.indexOf("last_refreshed") + 1;
  const textCol = GUEST_HEADERS.indexOf("live_status_text") + 1;
  sheet.getRange(rowIndex, tsCol).setValue(new Date().toISOString());
  sheet.getRange(rowIndex, textCol).setValue(live_status_text);
  return { success: true };
}

// ── Helper: find row number by group_id ───────────────────────
function findGuestRow(sheet, group_id) {
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === group_id) return i + 2;
  }
  return null;
}
