/**
 * Save-the-Date RSVP webhook for Anjali & Nisarga.
 *
 * Deploy: Open the Sheet (id below) → Extensions → Apps Script →
 * paste this in (replacing the default stub) → fill in TO_ADDRESSES →
 * Save → Run testWriteRow → grant OAuth → Deploy → New deployment →
 * Web app → Execute as: Me, Access: Anyone → copy /exec URL.
 *
 * Iterate later (preserves URL):
 *   Edit → Save → Deploy → Manage deployments → ✏️ → Version: New
 *   version → Deploy. Never click "New deployment" after the first
 *   one — that mints a new URL and breaks the form.
 *
 * The committed copy at site/apps-script/rsvp.gs keeps TO_ADDRESSES
 * as [TO_EMAIL_*] placeholders. The deployed copy in the Apps Script
 * editor must have real values. See _project-context/03-decisions-log.md.
 *
 * SCHEMA (5 columns): Timestamp | Name | Email | Response | Message
 *
 * Note: code duplication with ../wedding-website/apps-script/rsvp.gs
 * is intentional. Each project is small and self-contained;
 * factoring shared validation across two trivially-deployed scripts
 * is not worth the cost (see 03-decisions-log.md).
 */

const SHEET_ID     = '1Ki4Vh_O8VLmMG84gJpqW7CSBG0yJugWEHMw9MaYmcnI';
const SHEET_NAME   = 'RSVPs';
const TO_ADDRESSES = ['[TO_EMAIL_1]', '[TO_EMAIL_2]'];

function doPost(e) {
  try {
    const params = (e && e.parameter) || {};
    const t = v => (v == null ? '' : String(v).trim());

    const name  = t(params.name);
    const email = t(params.email);

    // Validate BEFORE touching the Sheet. Single generic error for any
    // failure so we don't leak which field was bad.
    if (!name || !email) {
      return _json({
        result: 'error',
        message: "Sorry, we couldn't record your response. Please double-check your details and try again.",
      });
    }

    const row = _buildRow_(params, name, email);
    _appendRow_(row);
    _notify_(row);
    return _json({ result: 'success' });

  } catch (err) {
    console.error('doPost failed:', err);
    return _json({
      result: 'error',
      message: 'Something went wrong. Please try again or email us directly.',
    });
  }
}

function doGet() {
  return _json({ result: 'error', message: 'POST only.' });
}

// Row order matches the Sheet header row:
// Timestamp | Name | Email | Response | Message
function _buildRow_(p, name, email) {
  const t = v => (v == null ? '' : String(v).trim());
  return [
    new Date().toISOString(),  // UTC, ISO 8601 with ms + 'Z'
    name,
    email,
    t(p.response),
    t(p.message),
  ];
}

// LockService serializes concurrent doPost calls so two submissions
// at the same instant don't trample each other on appendRow.
function _appendRow_(row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet tab "' + SHEET_NAME + '" not found in spreadsheet ' + SHEET_ID);
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

// Mail failures must not roll back the row — by this point it's saved.
function _notify_(row) {
  const [timestamp, name, email, response, message] = row;
  const subject = 'New Save-the-Date response from ' + name;
  const body = [
    'Name:           ' + name,
    'Email:          ' + email,
    'Response:       ' + (response || '(not specified)'),
    'Message:        ' + (message  || '(none)'),
    'Timestamp UTC:  ' + timestamp,
  ].join('\n');

  TO_ADDRESSES.forEach(function (addr) {
    if (!addr || addr.indexOf('[') === 0) return;  // skip unfilled tokens
    try {
      MailApp.sendEmail(addr, subject, body);
    } catch (mailErr) {
      console.error('MailApp.sendEmail failed for ' + addr + ':', mailErr);
    }
  });
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Smoke test — run from the Apps Script editor before deploying.
// Grants OAuth + verifies Sheet write + email side-effects. Writes a
// row tagged DELETE ME — delete the row from the Sheet after.
function testWriteRow() {
  const fakeEvent = { parameter: {
    name:     'TEST SUBMISSION — DELETE ME',
    email:    'test@example.com',
    response: 'Hoping to attend!',
    message:  'Smoke test from Apps Script editor',
  }};
  Logger.log(doPost(fakeEvent).getContent());
}
