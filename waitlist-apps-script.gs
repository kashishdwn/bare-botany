/**
 * Bare Botany — Waitlist capture (Google Apps Script)
 * Appends each website signup (email OR phone) as a row in your Google Sheet,
 * and tells the site whether the contact was newly added or already existed.
 *
 * The website calls this via JSONP (a GET with a ?callback= param) so the
 * browser can actually READ the reply — that's how we show "you're already
 * on the list" vs "thank you".
 *
 * ─────────────────────────────────────────────────────────────
 * AFTER EDITING THIS CODE you MUST redeploy for changes to go live:
 *   Deploy → Manage deployments → ✏️ (edit) → Version: "New version" → Deploy
 * The /exec URL stays the same.
 * ─────────────────────────────────────────────────────────────
 */

var SHEET_ID = '1OqvfN0kLK0KnYd0x670CwxgmC2I5CkJ5-le9Y7JLVbc';
var SHEET_NAME = ''; // '' = first tab

function getSheet_() {
  var ss = (SHEET_ID && SHEET_ID.indexOf('PASTE_') !== 0)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet — set SHEET_ID.');
  var sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
  if (!sheet) throw new Error('Sheet/tab not found.');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Type', 'Contact', 'Source', 'Page']);
  }
  return sheet;
}

// Shared handler for both GET (JSONP from the site) and POST.
function handle_(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();

    // The website validates email/phone (incl. per-country length) before
    // sending, so here we just trust the values and store them.
    var contact = (params.contact || params.email || params.phone || '').toString().trim();
    var type = (params.type || '').toString().trim().toLowerCase();
    if (!contact) return { ok: false, error: 'missing_contact' };

    // DUPLICATE CHECK — if this contact already exists, don't add again.
    var n = Math.max(sheet.getLastRow() - 1, 0);
    if (n > 0) {
      var existing = sheet.getRange(2, 3, n, 1).getValues(); // column C = Contact
      var lower = contact.toLowerCase();
      for (var i = 0; i < existing.length; i++) {
        if ((existing[i][0] || '').toString().trim().toLowerCase() === lower) {
          return { ok: true, duplicate: true, type: type };
        }
      }
    }

    sheet.appendRow([
      params.ts || new Date().toISOString(),
      type,
      contact,
      params.source || '',
      params.page || ''
    ]);
    return { ok: true, duplicate: false, type: type };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  // A plain visit (no contact) is a health check.
  if (!params.contact && !params.email && !params.phone) {
    var info;
    try {
      var sheet = getSheet_();
      info = { ok: true, service: 'Bare Botany waitlist', sheet: sheet.getName(),
               rows: Math.max(sheet.getLastRow() - 1, 0) };
    } catch (err) { info = { ok: false, error: String(err) }; }
    return reply_(info, params.callback);
  }
  return reply_(handle_(params), params.callback);
}

function doPost(e) {
  var params = {};
  try { params = JSON.parse(e.postData.contents); } catch (err) { params = {}; }
  return reply_(handle_(params), null);
}

// Returns JSONP if a callback name is given, otherwise plain JSON.
function reply_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback) {
    // allow only safe callback identifiers
    if (/^[A-Za-z_$][\w$]*$/.test(callback)) {
      return ContentService
        .createTextOutput(callback + '(' + body + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
