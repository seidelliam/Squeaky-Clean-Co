/**
 * Squeaky Clean Co — rental application intake.
 *
 * Form submit does two things at once:
 *   1. appends a row to the Applications tab
 *   2. emails the office the JSON record with the applicant's documents attached
 *
 * Applications are LEADS, not renters. Nothing touches the Renters tab until
 * you approve an application and run "Promote to Renters" from the menu — that
 * keeps unapproved leads out of your Dashboard counts, and keeps the form away
 * from the formula columns on Renters.
 *
 * doPost also handles two other posters:
 *   - click beacons from Messenger outreach links (api/click.js) log to the
 *     Clicks tab; see Squeaky Clean → Build ref report.
 *   - partner intake from the /vendor form (api/vendor.js) lands in the
 *     Vendor Leads tab with a plain email. It never touches the real
 *     Vendors tab — copy a lead across by hand when you onboard them.
 *
 * Deploy: Extensions > Apps Script, paste this in, fill CONFIG, run
 * setupApplicationsTab() once, then Deploy > New deployment > Web app
 * ("Execute as: Me", "Who has access: Anyone"). Put the /exec URL in Vercel.
 */

/* ===================== CONFIG ===================== */

var SHEET_ID     = '1qXDexIzX9wiROynCKTa-yL8EdxmDImIHziB-gzCjBg0';   // long id in the Sheet's URL
var NOTIFY_TO    = 'squeaky.clean.co.office@gmail.com';
var SHARED_TOKEN = 'b5784de3ef47a6f06aa738f346f5990884aa3c02234deb20f6911f9b05aadb30'; // must match Vercel's APPS_SCRIPT_TOKEN
var INTERNAL_KEY = 'f9e979869a435f4a086c39baee27de19c47a00631ef5c664'; // read-only dashboard passphrase; must match Vercel's INTERNAL_KEY

var APPLICATIONS_TAB = 'Applications';
var RENTERS_TAB      = 'Renters';
var CLICKS_TAB        = 'Clicks';
var REF_REPORT_TAB    = 'Ref Report';
var VENDOR_LEADS_TAB  = 'Vendor Leads';   // partner intake from /vendor — leads only, never the real Vendors tab

/** Renters puts its column names on row 2 — row 1 is the merged title banner. */
var RENTERS_HEADER_ROW = 2;
/** First data row on Renters. */
var RENTERS_FIRST_DATA_ROW = 3;

/** Defaults applied when promoting an application into Renters. */
var DEFAULT_MONTHLY_RATE = 80;
var DEFAULT_MY_CUT       = 10;
var NEW_RENTER_STATUS    = 'Pending Delivery';

/**
 * Renters columns this script is allowed to write, by header name.
 * Every column NOT listed here is a formula (Vendor Name, Within Radius?,
 * the Pmts counts, Trust Score, Trust Rating, Lifetime Paid) and is left
 * untouched so it keeps calculating.
 */
var RENTER_WRITABLE = [
  'Renter ID', 'Renter Name', 'Phone', 'Email', 'Street Address', 'City', 'Zip',
  'Hookup Type', 'Assigned Vendor ID', 'Distance to Vendor (mi)', 'Monthly Rate',
  'My Cut', 'Start Date', 'Status', 'Notes'
];

var APPLICATION_HEADERS = [
  'Application ID', 'Submitted At', 'Full Name', 'Phone', 'Email',
  'Preferred Language', 'City', 'Street Address', 'ZIP', 'Space Type',
  'Preferred Delivery Date', 'Lease Term (Months)', 'Monthly Rate (USD)',
  'Outlet Photo Attached', 'Consent Accepted', 'ID Check Acknowledged',
  'Review Status', 'Promoted to Renter ID',
  'Ref' // outreach attribution — see the click-tracking note at the top of this file
];

var VENDOR_LEAD_HEADERS = [
  'Submission ID', 'Submitted At', 'Name / Business', 'Phone', 'Email',
  'Primary City', 'Base ZIP', 'Travel / Radius', 'Services',
  'Notes & Requirements', 'Review Status'
];

/* ===================== ENDPOINT ===================== */

function doPost(e) {
  var body = null;
  try {
    body = JSON.parse(e.postData.contents);
    if (body.token !== SHARED_TOKEN) return reply_({ ok: false, error: 'unauthorized' });

    // Click beacon from api/click.js — logs to Clicks, never touches
    // Applications or sends mail. Kept out of the try/catch's mail-fallback
    // below since a click has no applicant to email anything to.
    if (body.click) return handleClick_(body.click);

    // Partner intake from api/vendor.js — its own leads tab and its own
    // (plain, no-attachment) email.
    if (body.vendor) return handleVendor_(body.vendor);

    var rec = body.record;
    if (!rec || !rec.applicationId) return reply_({ ok: false, error: 'bad_payload' });

    var row = appendApplication_(rec);
    sendMail_(rec);
    return reply_({ ok: true, applicationId: rec.applicationId, row: row });
  } catch (err) {
    // Get the operator the data even if the Sheet write failed.
    try { if (body && body.record) sendMail_(body.record, String(err)); } catch (ignored) {}
    return reply_({ ok: false, error: String(err) });
  }
}

function handleClick_(click) {
  try {
    appendClick_(click);
    return reply_({ ok: true });
  } catch (err) {
    // Best-effort logging — a Sheet hiccup here must never surface to the
    // visitor. api/click.js ignores this response either way.
    return reply_({ ok: false, error: String(err) });
  }
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===================== INTERNAL READ (doGet) ===================== */

/**
 * Read-only snapshot for the internal dashboard (site/internal, via
 * api/internal.js). Gated by INTERNAL_KEY passed as ?key=. Returns
 * Renters, Applications, and Vendor Leads as arrays of plain objects.
 * A bare visit with no key just gets {ok:false} — harmless.
 */
function doGet(e) {
  try {
    var key = (e && e.parameter && e.parameter.key) || '';
    if (key !== INTERNAL_KEY) return reply_({ ok: false, error: 'unauthorized' });

    return reply_({
      ok: true,
      generatedAt: new Date().toISOString(),
      renters:      readTab_(RENTERS_TAB,      RENTERS_HEADER_ROW, RENTERS_FIRST_DATA_ROW, 'Renter ID'),
      applications: readTab_(APPLICATIONS_TAB, 1, 2, 'Application ID'),
      vendorLeads:  readTab_(VENDOR_LEADS_TAB, 1, 2, 'Submission ID')
    });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  }
}

/**
 * One tab as an array of objects keyed by its header row. Rows whose
 * keyField cell is blank are skipped — Renters keeps its formulas filled
 * down past the real data. Dates come back as ISO strings so they survive
 * JSON.
 */
function readTab_(name, headerRow, firstDataRow, keyField) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!sh) return [];
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < firstDataRow || lastCol < 1) return [];

  var headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  var keyIdx = headers.indexOf(keyField);
  var rows = sh.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol).getValues();

  var out = [];
  rows.forEach(function (r) {
    if (keyIdx > -1 && String(r[keyIdx]).trim() === '') return;
    var obj = {};
    headers.forEach(function (h, i) {
      if (!h) return;
      var v = r[i];
      obj[h] = (v instanceof Date) ? v.toISOString() : v;
    });
    out.push(obj);
  });
  return out;
}

/* ===================== APPLICATIONS TAB ===================== */

function appendApplication_(rec) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(APPLICATIONS_TAB);
  if (!sh) throw new Error('No tab named "' + APPLICATIONS_TAB + '" — run setupApplicationsTab() once.');

  var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  var flat = rec.flatRow || {};
  var docs = (rec.documents || []).map(function (d) { return d.kind + ':' + d.fileName; }).join(' | ');

  var row = headers.map(function (h) {
    var key = String(h).trim();
    if (key === 'Review Status') return 'New';
    if (key === 'Promoted to Renter ID') return '';
    if (Object.prototype.hasOwnProperty.call(flat, key)) return flat[key];
    return '';
  });

  sh.appendRow(row);
  return sh.getLastRow();
}

/* ===================== EMAIL ===================== */

function sendMail_(rec, warning) {
  var flat = rec.flatRow || {};
  var docs = rec.documents || [];

  var lines = Object.keys(flat).map(function (k) {
    var v = flat[k];
    return k + ': ' + (v === null || v === undefined ? '' : v);
  });

  var body = 'New rental application\n\n' + lines.join('\n') +
    '\n\nDocuments attached: ' + docs.length +
    '\n\nThe full record is attached as JSON.';

  if (warning) {
    body = '!! The Applications row could NOT be written: ' + warning +
           '\n!! The data below is intact — add it by hand.\n\n' + body;
  }

  var attachments = [];
  var clean = JSON.parse(JSON.stringify(rec));
  clean.documents = docs.map(function (d, i) {
    return {
      kind: d.kind, fileName: d.fileName, mimeType: d.mimeType,
      byteSize: d.byteSize, attachedAs: attachmentName_(d, i)
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
  var m = String(d.mimeType || ''), ext = 'bin';
  if (m.indexOf('jpeg') > -1 || m.indexOf('jpg') > -1) ext = 'jpg';
  else if (m.indexOf('png') > -1) ext = 'png';
  else if (m.indexOf('pdf') > -1) ext = 'pdf';
  else if (m.indexOf('webp') > -1) ext = 'webp';
  return (d.kind || 'document') + '-' + (i + 1) + '.' + ext;
}

/* ===================== CLICKS TAB (outreach attribution) ===================== */

var CLICK_HEADERS = ['Ref', 'Clicked At', 'Landing Path', 'fbclid', 'User Agent'];

function appendClick_(click) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(CLICKS_TAB);
  if (!sh) sh = createClicksTab_();

  sh.appendRow([
    String((click && click.ref) || 'direct'),
    (click && click.ts) ? new Date(click.ts) : new Date(),
    String((click && click.landingPath) || ''),
    String((click && click.fbclid) || ''),
    String((click && click.ua) || '')
  ]);
}

function createClicksTab_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.insertSheet(CLICKS_TAB);
  sh.getRange(1, 1, 1, CLICK_HEADERS.length)
    .setValues([CLICK_HEADERS])
    .setFontWeight('bold')
    .setBackground('#2e5c9a')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, CLICK_HEADERS.length);
  return sh;
}

/** Run from the menu if you want the tab ready before the first real click
 *  arrives. Not required — appendClick_ creates it automatically. */
function setupClicksTab() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  if (!ss.getSheetByName(CLICKS_TAB)) createClicksTab_();
  SpreadsheetApp.getUi().alert('"' + CLICKS_TAB + '" is ready.');
}

/**
 * Squeaky Clean → Build ref report. Rebuilds the "Ref Report" tab from
 * scratch: one row per ref code seen in Clicks (plus any ref that only
 * shows up on an Application, in case a beacon ever failed to log), with
 * a click count, the first click time, and whether that ref converted —
 * so "clicked but never applied" is visible at a glance. To export,
 * open the tab and File → Download → CSV, same as any other sheet.
 */
function buildRefReport() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var clicksSh = ss.getSheetByName(CLICKS_TAB);
  var appsSh = ss.getSheetByName(APPLICATIONS_TAB);

  var clicks = (clicksSh && clicksSh.getLastRow() > 1)
    ? clicksSh.getRange(2, 1, clicksSh.getLastRow() - 1, 2).getValues()
    : [];

  var submittedRefs = {};
  if (appsSh) {
    var appHeaders = appsSh.getRange(1, 1, 1, appsSh.getLastColumn()).getValues()[0];
    var refCol = appHeaders.indexOf('Ref');
    var idCol = appHeaders.indexOf('Application ID');
    if (refCol > -1 && appsSh.getLastRow() > 1) {
      appsSh.getRange(2, 1, appsSh.getLastRow() - 1, appsSh.getLastColumn()).getValues().forEach(function (r) {
        var ref = String(r[refCol] || '').trim();
        if (!ref) return;
        (submittedRefs[ref] = submittedRefs[ref] || []).push(idCol > -1 ? r[idCol] : '');
      });
    }
  }

  var stats = {}; // ref -> { count, first }
  clicks.forEach(function (r) {
    var ref = String(r[0] || '').trim();
    if (!ref) return;
    var at = r[1];
    if (!stats[ref]) stats[ref] = { count: 0, first: at };
    stats[ref].count++;
    if (at && (!stats[ref].first || new Date(at) < new Date(stats[ref].first))) stats[ref].first = at;
  });
  Object.keys(submittedRefs).forEach(function (ref) {
    if (!stats[ref]) stats[ref] = { count: 0, first: '' };
  });

  var refs = Object.keys(stats).sort();
  var out = [['Ref', 'Clicks', 'First Click', 'Submitted?', 'Application ID(s)']];
  refs.forEach(function (ref) {
    var s = stats[ref], apps = submittedRefs[ref] || [];
    out.push([ref, s.count, s.first || '', apps.length ? 'Yes' : 'No', apps.join(', ')]);
  });

  var sh = ss.getSheetByName(REF_REPORT_TAB);
  if (!sh) sh = ss.insertSheet(REF_REPORT_TAB);
  sh.clear();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(1, 1, 1, out[0].length)
    .setFontWeight('bold').setBackground('#2e5c9a').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, out[0].length);

  SpreadsheetApp.getUi().alert('"' + REF_REPORT_TAB + '" rebuilt: ' + refs.length + ' ref code(s).');
}

/* ===================== VENDOR LEADS (partner intake) ===================== */

function handleVendor_(v) {
  try {
    if (!v || !v.submissionId) return reply_({ ok: false, error: 'bad_payload' });
    var row = appendVendorLead_(v);
    sendVendorMail_(v);
    return reply_({ ok: true, submissionId: v.submissionId, row: row });
  } catch (err) {
    // Get the operator the data even if the Sheet write failed.
    try { if (v) sendVendorMail_(v, String(err)); } catch (ignored) {}
    return reply_({ ok: false, error: String(err) });
  }
}

function appendVendorLead_(v) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(VENDOR_LEADS_TAB) || createVendorLeadsTab_();

  var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  var flat = v.flatRow || {};
  var row = headers.map(function (h) {
    var key = String(h).trim();
    if (key === 'Review Status') return 'New';
    if (Object.prototype.hasOwnProperty.call(flat, key)) return flat[key];
    return '';
  });

  sh.appendRow(row);
  return sh.getLastRow();
}

function createVendorLeadsTab_() {
  var sh = SpreadsheetApp.openById(SHEET_ID).insertSheet(VENDOR_LEADS_TAB);
  sh.getRange(1, 1, 1, VENDOR_LEAD_HEADERS.length)
    .setValues([VENDOR_LEAD_HEADERS])
    .setFontWeight('bold')
    .setBackground('#2e5c9a')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, VENDOR_LEAD_HEADERS.length);
  return sh;
}

/** Run from the menu if you want the tab ready first. Not required —
 *  appendVendorLead_ creates it automatically on the first submission. */
function setupVendorLeadsTab() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  if (!ss.getSheetByName(VENDOR_LEADS_TAB)) createVendorLeadsTab_();
  SpreadsheetApp.getUi().alert('"' + VENDOR_LEADS_TAB + '" is ready.');
}

function sendVendorMail_(v, warning) {
  var flat = v.flatRow || {};
  var lines = Object.keys(flat).map(function (k) {
    return k + ': ' + (flat[k] === null || flat[k] === undefined ? '' : flat[k]);
  });
  var body = 'New vendor / partner inquiry\n\n' + lines.join('\n');
  if (warning) {
    body = '!! The Vendor Leads row could NOT be written: ' + warning +
           '\n!! The data below is intact — add it by hand.\n\n' + body;
  }
  MailApp.sendEmail({
    to: NOTIFY_TO,
    replyTo: flat['Email'] || NOTIFY_TO,
    subject: 'Vendor inquiry — ' + (flat['Name / Business'] || 'unknown') +
             ' (' + (flat['Primary City'] || '?') + ')',
    body: body
  });
}

/* ===================== PROMOTE TO RENTERS ===================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Squeaky Clean')
    .addItem('Promote selected application to Renters', 'promoteSelectedApplication')
    .addSeparator()
    .addItem('Build ref report (clicks vs. applications)', 'buildRefReport')
    .addSeparator()
    .addItem('Set up Applications tab', 'setupApplicationsTab')
    .addItem('Set up Clicks tab', 'setupClicksTab')
    .addItem('Set up Vendor Leads tab', 'setupVendorLeadsTab')
    .addItem('Test intake (row + email)', 'testEndToEnd')
    .addToUi();
}

/**
 * Select any cell in an Applications row, then run this. It writes a new
 * Renters row using ONLY the input columns — the formula columns in that row
 * are left exactly as they are, so they keep calculating.
 */
function promoteSelectedApplication() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var apps = ss.getSheetByName(APPLICATIONS_TAB);
  var active = ss.getActiveSheet();

  if (!apps || active.getName() !== APPLICATIONS_TAB) {
    ui.alert('Open the ' + APPLICATIONS_TAB + ' tab and select the row you want to promote.');
    return;
  }

  var rowNum = active.getActiveRange().getRow();
  if (rowNum < 2) { ui.alert('Select an application row, not the header.'); return; }

  var headers = apps.getRange(1, 1, 1, apps.getLastColumn()).getValues()[0];
  var values = apps.getRange(rowNum, 1, 1, apps.getLastColumn()).getValues()[0];
  var app = {};
  headers.forEach(function (h, i) { app[String(h).trim()] = values[i]; });

  if (!app['Application ID']) { ui.alert('That row has no Application ID.'); return; }
  if (app['Promoted to Renter ID']) {
    ui.alert('Already promoted as ' + app['Promoted to Renter ID'] + '.');
    return;
  }

  // No documents are uploaded anymore. The renter shows a driver's license
  // or state ID in person at delivery; the form only records that they
  // agreed to that (the "ID Check Acknowledged" column), and that flag
  // carries into the renter's Notes below. If it somehow reads anything
  // other than "Yes", warn before a machine goes out.
  if (String(app['ID Check Acknowledged']).toLowerCase() !== 'yes') {
    var resp = ui.alert(
      'ID check not acknowledged',
      'This application does not have the delivery ID check acknowledged. ' +
      'Promote to Renters anyway?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  }

  var renterId = writeRenterRow_(app);

  var promotedCol = headers.indexOf('Promoted to Renter ID') + 1;
  var statusCol = headers.indexOf('Review Status') + 1;
  if (promotedCol > 0) apps.getRange(rowNum, promotedCol).setValue(renterId);
  if (statusCol > 0) apps.getRange(rowNum, statusCol).setValue('Approved');

  ui.alert('Added ' + renterId + ' to ' + RENTERS_TAB + '.\n\n' +
           'Verify the driver\'s license / ID at delivery.\n' +
           'Still to fill in by hand: Assigned Vendor ID, Distance to Vendor (mi), My Cut.');
}

function writeRenterRow_(app) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(RENTERS_TAB);
  if (!sh) throw new Error('No tab named "' + RENTERS_TAB + '"');

  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(RENTERS_HEADER_ROW, 1, 1, lastCol).getValues()[0];

  var targetRow = firstBlankRenterRow_(sh);
  var renterId = nextRenterId_(sh, headers);

  var out = {
    'Renter ID': renterId,
    'Renter Name': app['Full Name'],
    'Phone': app['Phone'],
    'Email': app['Email'],
    'Street Address': app['Street Address'],
    'City': app['City'],
    'Zip': app['ZIP'],
    'Hookup Type': 'Electric',
    'Monthly Rate': DEFAULT_MONTHLY_RATE,
    'My Cut': DEFAULT_MY_CUT,
    'Start Date': app['Preferred Delivery Date'] || '',
    'Status': NEW_RENTER_STATUS,
    'Notes': 'From ' + app['Application ID'] +
             ' · ' + (app['Lease Term (Months)'] || '?') + ' mo term' +
             ' · space ' + (app['Space Type'] || '?') +
             ' · lang ' + (app['Preferred Language'] || '') +
             ' · ID check ' +
               (String(app['ID Check Acknowledged']).toLowerCase() === 'yes' ? 'acknowledged, verify at delivery' : 'NOT acknowledged')
  };

  // Write cell by cell, and ONLY into columns on the writable list. Anything
  // else on this row is a formula and must be left alone.
  headers.forEach(function (h, i) {
    var key = String(h).trim();
    if (RENTER_WRITABLE.indexOf(key) === -1) return;
    if (!Object.prototype.hasOwnProperty.call(out, key)) return;
    sh.getRange(targetRow, i + 1).setValue(out[key]);
  });

  return renterId;
}

/**
 * First row whose Renter ID cell is empty. Renters has its formulas already
 * filled down past the data, so appendRow() would land BELOW them — this finds
 * the real next slot instead.
 */
function firstBlankRenterRow_(sh) {
  var maxRow = Math.max(sh.getMaxRows(), RENTERS_FIRST_DATA_ROW);
  var n = maxRow - RENTERS_FIRST_DATA_ROW + 1;
  var ids = sh.getRange(RENTERS_FIRST_DATA_ROW, 1, n, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === '') return RENTERS_FIRST_DATA_ROW + i;
  }
  return maxRow + 1;
}

function nextRenterId_(sh, headers) {
  var col = headers.indexOf('Renter ID') + 1;
  if (col < 1) col = 1;
  var n = Math.max(1, sh.getMaxRows() - RENTERS_FIRST_DATA_ROW + 1);
  var ids = sh.getRange(RENTERS_FIRST_DATA_ROW, col, n, 1).getValues();
  var max = 0;
  ids.forEach(function (r) {
    var m = String(r[0]).match(/^R-(\d+)/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return 'R-' + String(max + 1).padStart(3, '0');
}

/* ===================== SETUP HELPERS ===================== */

/** Run once. Creates the Applications tab with the headers the form writes. */
function setupApplicationsTab() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(APPLICATIONS_TAB);
  if (!sh) sh = ss.insertSheet(APPLICATIONS_TAB);

  sh.getRange(1, 1, 1, APPLICATION_HEADERS.length)
    .setValues([APPLICATION_HEADERS])
    .setFontWeight('bold')
    .setBackground('#2e5c9a')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, APPLICATION_HEADERS.length);
  SpreadsheetApp.getUi().alert('"' + APPLICATIONS_TAB + '" is ready.');
}

/** Run once to confirm the Sheet write and the email both work. */
function testEndToEnd() {
  var now = new Date().toISOString();
  var rec = {
    applicationId: 'SCC-TEST-0001',
    submittedAt: now,
    applicant: { fullName: 'Test Applicant', email: NOTIFY_TO },
    delivery: { city: 'Houston' },
    documents: [],
    flatRow: {
      'Application ID': 'SCC-TEST-0001', 'Submitted At': now,
      'Full Name': 'Test Applicant', 'Phone': '(555) 555-0123',
      'Email': NOTIFY_TO, 'City': 'Houston', 'ZIP': '77520',
      'Street Address': '1 Test St', 'Lease Term (Months)': 6,
      'Monthly Rate (USD)': 80, 'Consent Accepted': 'Yes',
      'ID Check Acknowledged': 'Yes', 'Ref': 'test-0000-xxxx'
    }
  };
  appendApplication_(rec);
  sendMail_(rec);
}
