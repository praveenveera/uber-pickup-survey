/**
 * Backend for the Pickup Story survey app.
 * Paste this into a Google Sheet's Apps Script editor (Extensions > Apps Script),
 * then deploy as a Web App. See DEPLOY.md in this folder for exact steps.
 *
 * Every submission becomes one row in a sheet called "Responses"
 * (created automatically on first submission).
 */

var SHEET_NAME = "Responses";

var COLUMNS = [
  "timestamp", "language", "role", "city", "cityOther", "scope4wheeler",
  "difficulty", "story", "location", "cause", "actions", "actionsOther",
  "timeLostMinutes", "helped", "repeatOccurrence", "more", "followupOk", "contact"
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getOrCreateSheet();
    var data = JSON.parse(e.postData.contents);
    var row = COLUMNS.map(function (key) {
      var value = data[key];
      return value === undefined || value === null ? "" : value;
    });
    sheet.appendRow(row);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Run this once manually from the Apps Script editor to confirm the sheet
 *  and header row are created before you deploy — optional, just a sanity check. */
function setupCheck() {
  getOrCreateSheet();
}
