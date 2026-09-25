/**
 * Caseboard daily email reminders — Google Apps Script
 *
 * SETUP (one time):
 * 1. Open the Caseboard Google Sheet (the one with the "Puzzles" tab).
 * 2. Extensions > Apps Script. Delete any starter code and paste this whole file in.
 * 3. Deploy > New deployment > gear icon > "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy, authorize it (it's your own script, asking for permission to send
 *    email and edit this sheet), then copy the "Web app URL" it gives you — it ends
 *    in /exec.
 * 4. Paste that URL into index.html as EMAIL_SIGNUP_ENDPOINT (replacing the
 *    'YOUR-APPS-SCRIPT-WEB-APP-URL' placeholder), then republish/push the site.
 * 5. Back in the Apps Script editor: click the clock icon (Triggers) > Add Trigger.
 *      Function: sendDailyReminders
 *      Event source: Time-driven
 *      Type: Day timer, pick a window like 8am-9am (Eastern) — after the midnight
 *      Eastern daily reset, so the day's case is definitely already live.
 *    Save, authorize again if asked.
 *
 * That's the whole setup — no server, no hosting, just this script living inside the
 * Sheet you already use for puzzle content.
 *
 * NOTES / LIMITS (worth knowing, not urgent to fix):
 * - MailApp.sendEmail has a daily quota (~100/day on a plain Gmail account, much
 *   higher on Google Workspace). Fine for a personal project; revisit if the list
 *   ever gets large.
 * - There's no unsubscribe link yet — if someone asks to be removed, delete their row
 *   from the Subscribers tab, or set their Status to "unsubscribed" (the send loop
 *   skips that value). Worth adding a real unsubscribe link later if the list grows.
 * - doPost has no auth/rate-limiting beyond "must look like an email" — fine for a
 *   low-traffic personal site, not something to expose more broadly as-is.
 */

var SUBSCRIBERS_TAB = 'Subscribers';
var PUZZLES_TAB = 'Puzzles';
var SITE_URL = 'https://jtoeman.github.io/caseboard/';

// Handles the POST from the site's email signup form. Body is a plain-text JSON
// string like {"email":"someone@example.com"} — sent with fetch's 'no-cors' mode from
// the browser, so there's no response the page can read; this just needs to persist
// the signup.
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var email = (body.email || '').trim().toLowerCase();
    if (!isValidEmail_(email)) {
      return ContentService.createTextOutput('invalid email');
    }
    var sheet = getOrCreateSubscribersSheet_();
    var existing = findSubscriberRow_(sheet, email);
    if (existing) {
      // Re-signing up un-does a prior unsubscribe rather than creating a duplicate row.
      sheet.getRange(existing, 3).setValue('active');
    } else {
      sheet.appendRow([email, new Date(), 'active']);
    }
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.message);
  }
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function getOrCreateSubscribersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUBSCRIBERS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(SUBSCRIBERS_TAB);
    sheet.appendRow(['Email', 'Signed up', 'Status']);
  }
  return sheet;
}

function findSubscriberRow_(sheet, email) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === email) return i + 1; // 1-indexed row
  }
  return null;
}

// Run daily by a time-driven trigger (see setup notes above). Emails every active
// subscriber a short note once a case is confirmed live for today.
function sendDailyReminders() {
  var todayStr = todayEasternStr_();
  var puzzleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PUZZLES_TAB);
  if (!puzzleSheet) return;

  var rows = puzzleSheet.getDataRange().getValues();
  var dates = [];
  var hasToday = false;
  for (var i = 1; i < rows.length; i++) { // skip header row
    var d = normalizeDate_(rows[i][0]);
    if (!d) continue;
    if (dates.indexOf(d) === -1) dates.push(d);
    if (d === todayStr) hasToday = true;
  }
  if (!hasToday) return; // nothing published for today yet — skip silently, no email sent
  dates.sort();
  var caseNumber = dates.indexOf(todayStr) + 1;

  var subsSheet = getOrCreateSubscribersSheet_();
  var subs = subsSheet.getDataRange().getValues();
  var subject = 'Caseboard — Case ' + caseNumber + ' is open';
  var body = 'Today’s case is up.\n\n' + SITE_URL + '\n\n' +
    'Keep your streak going!\n\n' +
    '(To stop these emails, reply to this message and let Jeremy know.)';

  for (var j = 1; j < subs.length; j++) {
    var email = subs[j][0];
    var status = String(subs[j][2] || '').toLowerCase();
    if (!email || status === 'unsubscribed') continue;
    try {
      MailApp.sendEmail(email, subject, body);
    } catch (err) {
      // One bad address shouldn't stop the rest of the send.
      Logger.log('Failed to email ' + email + ': ' + err.message);
    }
  }
}

// Mirrors the site's own date normalization (gviz-style "M/D/YYYY" or a real Date
// cell) into "YYYY-MM-DD" so it compares correctly against todayEasternStr_().
function normalizeDate_(cell) {
  if (Object.prototype.toString.call(cell) === '[object Date]') {
    return Utilities.formatDate(cell, 'America/New_York', 'yyyy-MM-dd');
  }
  var s = String(cell || '').trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return s || null;
  var mm = m[1].length < 2 ? '0' + m[1] : m[1];
  var dd = m[2].length < 2 ? '0' + m[2] : m[2];
  return m[3] + '-' + mm + '-' + dd;
}

function todayEasternStr_() {
  return Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
}
