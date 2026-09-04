/**
 * Squeaky Clean Co — rental application intake.
 *
 * One endpoint, two outputs: appends a row to the Google Sheet, and emails
 * the office a JSON copy with the applicant's documents attached.
 *
 * Deploy: Extensions > Apps Script, paste this in, fill the CONFIG block,
 * then Deploy > New deployment > Web app, "Execute as: Me", "Who has
 * access: Anyone". Copy the /exec URL into Vercel as APPS_SCRIPT_URL.
 */

/* ================= CONFIG ================= */

var SHEET_ID    = 'PASTE_SHEET_ID_HERE';   // the long id in the Sheet's URL
var SHEET_NAME  = 'Applications';          // the tab to append to
var NOTIFY_TO   = 'squeaky.clean.co.office@gmail.com';
var SHARED_TOKEN = 'PASTE_A_LONG_RANDOM_STRING_HERE'; // must match Vercel's APPS_SCRIPT_TOKEN

/**
 * Only needed when a Sheet column header differs from the payload's own key.
 * Left side = the header text in row 1 of your Sheet.
 * Right side = the key in the application's flatRow.
 * Headers that already match a flatRow key need no entry here.
 */
var ALIASES = {
  // 'Date Received' : 'Submitted At',
  // 'Customer Name' : 'Full Name',
  // 'Term'          : 'Lease Term (Months)',
};

/* ================= ENDPOINT ================= */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.token !== SHARED_TOKEN) return reply_({ ok: false, error: 'unauthorized' });

    var rec = body.record;
    if (!rec || !rec.applicationId) return reply_({ ok: false, error: 'bad_payload' });

    var rowNumber = appendRow_(rec);
    sendMail_(rec);

    return reply_({ ok: true, applicationId: rec.applicationId, row: rowNumber });
  } catch (err) {
    // Still try to get the operator the data even if the Sheet write failed.
    try { if (body && body.record) sendMail_(body.record, String(err)); } catch (ignored) {}
    return reply_({ ok: false, error: String(err) });
  }
}

function reply_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= SHEET ================= */

/**
 * Appends one row, matching values to the Sheet's own header row by NAME.
 * Reordering columns, inserting new ones, or renaming via ALIASES all keep
 * working without touching this script.
 */
function appendRow_(rec) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('No sheet tab named "' + SHEET_NAME + '"');

  var lastCol = Math.max(1, sh.getLastColumn());
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var flat = rec.flatRow || {};

  var row = headers.map(function (h) {
    var key = String(h).trim();
    if (!key) return '';
    if (Object.prototype.hasOwnProperty.call(flat, key)) return flat[key];
    var alias = ALIASES[key];
    if (alias && Object.prototype.hasOwnProperty.call(flat, alias)) return flat[alias];
    return '';
  });

  sh.appendRow(row);
  return sh.getLastRow();
}

/* ================= EMAIL ================= */

function sendMail_(rec, warning) {
  var flat = rec.flatRow || {};
  var docs = rec.documents || [];

  var lines = Object.keys(flat).map(function (k) {
    return k + ': ' + (flat[k] === null || flat[k] === undefined ? '' : flat[k]);
  });

  var body = 'New rental application\n\n' + lines.join('\n') +
    '\n\nDocuments attached: ' + (docs.length || 0) +
    '\n\nThe full record is attached as JSON.';

  if (warning) {
    body = '!! The Sheet row could NOT be written: ' + warning +
           '\n!! The application data below is intact — add it by hand.\n\n' + body;
  }

  var attachments = [];

  // Clean JSON — the base64 blobs are stripped since the files ride along
  // as real attachments; keeps the JSON readable and the email small.
  var clean = JSON.parse(JSON.stringify(rec));
  clean.documents = docs.map(function (d, i) {
    return {
      kind: d.kind,
      fileName: d.fileName,
      mimeType: d.mimeType,
      byteSize: d.byteSize,
      attachedAs: attachmentName_(d, i)
    };
  });
  attachments.push(Utilities.newBlob(
    JSON.stringify(clean, null, 2), 'application/json', rec.applicationId + '.json'
  ));

  docs.forEach(function (d, i) {
    var b64 = String(d.dataUrl || '').split(',')[1];
    if (!b64) return;
    try {
      attachments.push(Utilities.newBlob(
        Utilities.base64Decode(b64), d.mimeType || 'application/octet-stream', attachmentName_(d, i)
      ));
    } catch (err) { /* skip an unreadable attachment rather than lose the email */ }
  });

  MailApp.sendEmail({
    to: NOTIFY_TO,
    replyTo: (rec.applicant && rec.applicant.email) || NOTIFY_TO,
    subject: 'Application ' + rec.applicationId + ' — ' +
             ((rec.applicant && rec.applicant.fullName) || 'unknown') + ' (' +
             ((rec.delivery && rec.delivery.city) || '?') + ')',
    body: body,
    attachments: attachments
  });
}

function attachmentName_(d, i) {
  var ext = 'bin';
  var m = String(d.mimeType || '');
  if (m.indexOf('jpeg') > -1 || m.indexOf('jpg') > -1) ext = 'jpg';
  else if (m.indexOf('png') > -1) ext = 'png';
  else if (m.indexOf('pdf') > -1) ext = 'pdf';
  else if (m.indexOf('webp') > -1) ext = 'webp';
  return (d.kind || 'document') + '-' + (i + 1) + '.' + ext;
}

/* ================= SETUP HELPERS ================= */

/** Run once from the editor to write the header row this script expects. */
function writeDefaultHeaders() {
  var headers = [
    'Application ID', 'Submitted At', 'Full Name', 'Phone', 'Email',
    'Preferred Language', 'City', 'Street Address', 'ZIP', 'Space Type',
    'Preferred Delivery Date', 'Lease Term (Months)', 'Monthly Rate (USD)',
    'Employer or Income Source', 'Monthly Income (USD)', 'Paystubs Attached',
    'License Images Attached', 'Outlet Photo Attached', 'Consent Accepted'
  ];
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/** Run once to confirm the Sheet write and the email both work end to end. */
function testEndToEnd() {
  var now = new Date().toISOString();
  sendMail_({
    applicationId: 'SCC-TEST-0001',
    submittedAt: now,
    applicant: { fullName: 'Test Applicant', email: NOTIFY_TO },
    delivery: { city: 'Houston' },
    documents: [],
    flatRow: { 'Application ID': 'SCC-TEST-0001', 'Submitted At': now, 'Full Name': 'Test Applicant' }
  });
  appendRow_({
    applicationId: 'SCC-TEST-0001',
    flatRow: { 'Application ID': 'SCC-TEST-0001', 'Submitted At': now, 'Full Name': 'Test Applicant' }
  });
}
