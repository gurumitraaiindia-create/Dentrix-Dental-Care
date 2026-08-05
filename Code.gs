/**
 * Dentrix Dental Care — Appointment Booking
 * Google Apps Script Web App backend.
 *
 * Handles:
 *   - doGet(e)  : health check for the deployed web app
 *   - doPost(e) : saves { name, email, phone } into a Google Sheet
 *                 named "Responses" (created on first use with bold headers)
 *
 * Columns stored: Timestamp | Name | Email | Phone
 */

var SHEET_NAME = 'Responses';
var HEADERS = ['Timestamp', 'Name', 'Email', 'Phone'];

// Paste your Google Sheet ID here (the long string in the sheet's URL)
// OR bind the script to the spreadsheet and leave this empty.
var SHEET_ID = '';

/**
 * GET handler.
 * Returns a JSON status so you can confirm the web app is reachable.
 */
function doGet(e) {
  return jsonResponse({ status: 'success', message: 'Dentrix Dental Care Web App is running.' });
}

/**
 * POST handler.
 * Reads name, email and phone from the request body (JSON or form-encoded)
 * and appends a new row to the "Responses" sheet.
 */
function doPost(e) {
  try {
    var data = parsePayload(e);

    var name = cleanString(data.name);
    var email = cleanString(data.email);
    var phone = cleanString(data.phone);

    if (!name || !email || !phone) {
      return jsonResponse({ status: 'error', message: 'All fields are required.' });
    }

    var sheet = getResponsesSheet();

    sheet.appendRow([
      new Date(),
      name,
      email,
      phone
    ]);

    return jsonResponse({
      status: 'success',
      message: 'Your appointment request has been received.'
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Something went wrong: ' + err });
  }
}

/**
 * Parses the incoming request body.
 * Works with both JSON and application/x-www-form-urlencoded payloads.
 */
function parsePayload(e) {
  var raw = '';
  if (e && e.postData && e.postData.contents) {
    raw = e.postData.contents;
  }

  var result = {};

  if (raw) {
    var trimmed = raw.trim();

    if (trimmed.charAt(0) === '{') {
      result = JSON.parse(trimmed);
    } else {
      var params = trimmed.split('&');
      for (var i = 0; i < params.length; i++) {
        var pair = params[i].split('=');
        if (pair.length === 2) {
          result[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1].replace(/\+/g, ' '));
        }
      }
    }
  }

  if (e && e.parameter) {
    for (var key in e.parameter) {
      if (e.parameter.hasOwnProperty(key)) {
        result[key] = e.parameter[key];
      }
    }
  }

  return result;
}

/**
 * Returns the "Responses" sheet, creating it with bold headers if needed.
 */
function getResponsesSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Resolves the spreadsheet:
 *   1. A spreadsheet bound to the script (created via Extensions > Apps Script).
 *   2. The spreadsheet opened from SHEET_ID.
 * Throws a clear error if neither is available.
 */
function getSpreadsheet() {
  var ss = null;

  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    ss = null;
  }

  if (!ss && SHEET_ID) {
    ss = SpreadsheetApp.openById(SHEET_ID);
  }

  if (!ss) {
    throw new Error('Could not locate the spreadsheet. Set SHEET_ID in Code.gs or bind this script to a spreadsheet.');
  }

  return ss;
}

/**
 * Trims and normalizes an incoming string value.
 */
function cleanString(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

/**
 * Builds a JSON output.
 */
function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
