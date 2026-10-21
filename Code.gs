/**
 * Code.gs - Google Apps Script backend (updated)
 * Saves enquiry / appointment requests to the "Responses" tab.
 *
 * Accepts GET and POST. Fields: name, phone, enquiryType, service, message
 * Required: name, phone
 *
 * After pasting this, redeploy:
 * Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy
 * The web app URL stays the same.
 */

const SHEET_NAME = 'Responses';
const HEADERS = ['Timestamp', 'Name', 'Phone', 'Enquiry Type', 'Service Required', 'Message'];

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  const lock = LockService.getScriptLock();
  try {
    const data = extractData_(e);

    const name = clean_(data.name);
    const phone = clean_(data.phone);
    const enquiryType = clean_(data.enquiryType);
    const service = clean_(data.service);
    const message = clean_(data.message);

    if (!name || !phone) {
      return jsonResponse_({ success: false, error: 'Name and phone are required.' });
    }

    lock.waitLock(10000);

    const sheet = getOrCreateSheet_();
    sheet.appendRow([new Date(), name, phone, enquiryType, service, message]);

    return jsonResponse_({ success: true });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function extractData_(e) {
  let data = {};

  if (e && e.parameter) {
    data = Object.assign({}, e.parameter);
  }

  if (e && e.postData && e.postData.contents) {
    const type = (e.postData.type || '').toLowerCase();
    if (type.indexOf('application/json') !== -1) {
      try {
        data = Object.assign(data, JSON.parse(e.postData.contents));
      } catch (_) {}
    }
  }

  return data;
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function clean_(value) {
  if (value === undefined || value === null) return '';
  let str = String(value).trim();
  if (/^[=+\-@]/.test(str)) {
    str = "'" + str;
  }
  return str.substring(0, 1000);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
