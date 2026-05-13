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
    const upsert = _upsertRow_(row);
    _notify_(row, upsert);
    return _json({
      result: 'success',
      isUpdate: !upsert.isNew,
    });

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

// Upsert by email. If a row already exists with the same (trimmed,
// lowercased) email, overwrite it in place — keeps one row per
// submitter so guest counts don't double when someone resubmits a
// month later because they forgot they already responded.
// Position-in-sheet stays the same on update, so row order = order
// of first contact.
//
// LockService serializes concurrent invocations so two submissions
// at the same instant don't trample each other on the read-modify-
// write.
//
// Returns:
//   { isNew: true }                    — appended a new row
//   { isNew: false, previous: row }    — updated an existing row;
//                                        `previous` is the row's
//                                        prior contents (for change
//                                        detection in _notify_)
function _upsertRow_(row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet tab "' + SHEET_NAME + '" not found in spreadsheet ' + SHEET_ID);

    const data = sheet.getDataRange().getValues();
    const newEmail = String(row[2] || '').trim().toLowerCase();

    // Scan all rows except the header (data[0]) for an existing
    // submitter with the same email.
    for (let i = 1; i < data.length; i++) {
      const existingEmail = String(data[i][2] || '').trim().toLowerCase();
      if (existingEmail && existingEmail === newEmail) {
        // Sheet rows are 1-indexed; header is row 1, so data[i] sits
        // at sheet row i + 1.
        const sheetRowNumber = i + 1;
        const previous = data[i];
        sheet.getRange(sheetRowNumber, 1, 1, row.length).setValues([row]);
        return { isNew: false, previous: previous };
      }
    }

    // No match — append.
    sheet.appendRow(row);
    return { isNew: true };
  } finally {
    lock.releaseLock();
  }
}

// Mail failures must not roll back the row — by this point it's saved.
// Sends two kinds of mail per submission, with branching for the
// upsert-detected-duplicate case:
//   1. Couple alert: sent for NEW submissions and for UPDATES that
//      actually changed response or message. Suppressed for "duplicate
//      with no change" (the "they forgot they responded" case) to
//      keep the couple's inbox sane.
//   2. Submitter confirmation: always sent. Wording adjusts to
//      acknowledge updates so resubmitters know the system saw their
//      first attempt.
function _notify_(row, upsert) {
  const [timestamp, name, email, response, message] = row;
  const isUpdate = upsert && upsert.isNew === false;
  const prev = isUpdate ? upsert.previous : null;
  const prevResponse = prev ? String(prev[3] || '').trim() : '';
  const prevMessage  = prev ? String(prev[4] || '').trim() : '';
  const changed = isUpdate && (response !== prevResponse || message !== prevMessage);

  // 1. Couple alert — suppressed when an update doesn't actually change anything.
  const shouldAlert = !isUpdate || changed;
  if (shouldAlert) {
    const alertSubject = (isUpdate ? '[Updated] ' : '') + 'New Save-the-Date response from ' + name;
    const alertBodyLines = [
      isUpdate ? '(UPDATE — same email previously responded.)' : '',
      isUpdate ? '' : null,
      'Name:           ' + name,
      'Email:          ' + email,
      'Response:       ' + (response || '(not specified)'),
      'Message:        ' + (message  || '(none)'),
      'Timestamp UTC:  ' + timestamp,
    ];
    if (isUpdate) {
      alertBodyLines.push('', '--- Previous ---',
        'Response:       ' + (prevResponse || '(not specified)'),
        'Message:        ' + (prevMessage  || '(none)'));
    }
    const alertBody = alertBodyLines.filter(s => s !== null).join('\n');
    TO_ADDRESSES.forEach(function (addr) {
      if (!addr || addr.indexOf('[') === 0) return;
      try {
        MailApp.sendEmail(addr, alertSubject, alertBody);
      } catch (mailErr) {
        console.error('Couple alert failed for ' + addr + ':', mailErr);
      }
    });
  }

  // 2. Submitter confirmation — always sent; wording varies on update.
  const confirmSubject = isUpdate
    ? "We've updated your Save-the-Date response — Anjali & Nisarga"
    : "Thanks for responding to our Save the Date — Anjali & Nisarga";
  const intro = isUpdate
    ? "Thanks! We've updated your response. (Looks like you may have responded before — that's totally fine, we just kept the latest.)"
    : "Thanks for letting us know you got our save the date! We've recorded your response:";
  const confirmBody = [
    'Hi ' + name + ',',
    '',
    intro,
    '',
    '  Your response: ' + (response || '(no specific response chosen)'),
    (message ? '  Your note:     ' + message + '\n' : ''),
    'Save the date — June 12, 2027 in Chattanooga, TN.',
    'Formal invitation to follow.',
    '',
    'With love,',
    'Anjali & Nisarga',
  ].join('\n');
  try {
    MailApp.sendEmail(email, confirmSubject, confirmBody);
  } catch (mailErr) {
    console.error('Submitter confirmation failed for ' + email + ':', mailErr);
  }
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
    response: 'Planning to attend!',
    message:  'Smoke test from Apps Script editor',
  }};
  Logger.log(doPost(fakeEvent).getContent());
}
