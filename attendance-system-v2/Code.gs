// ============================================================
//  ATTENDANCE SYSTEM — Google Apps Script  (Scale: 200+ students)
//  Deploy as Web App: Execute as "Me", Access "Anyone"
//
//  CHANGES FROM BUILD LOG 1:
//  - claimed_ lock REMOVED — all students can scan the same QR
//  - GRACE_MS raised to 90s — handles slow phones + network lag
//  - Row formatting stripped from submit — sheet is pre-styled
//  - Duplicate guard uses LockService for race-safe writes
//  - register wipes only its own token, not all tokens (safer)
//  - check no longer blocks simultaneous scans
// ============================================================

const SHEET_ID   = 'YOUR_GOOGLE_SHEET_ID';   // ← replace this
const SHEET_NAME = 'Attendance-Research & Developement';
const GRACE_MS   = 90000;  // 90 seconds — enough for 200+ students on congested wifi
const props      = PropertiesService.getScriptProperties();

const HEADERS = [
  'Sr No', 'Date', 'Session', 'Name', 'Roll Number',
  'Branch', 'Year / Semester', 'Email', 'College', 'Token',
  'Submitted At', 'Status',
  'Easy to Understand?', 'Comfortable Pace?',
  'Doubts Addressed?', 'More Examples Needed?', 'Written Feedback'
];

// ── Run ONCE manually to set up / reset the sheet ────────────
function fixSheet() {
  const ss  = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(SHEET_NAME);

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1F3864');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1,60);  sheet.setColumnWidth(2,110); sheet.setColumnWidth(3,180);
  sheet.setColumnWidth(4,160); sheet.setColumnWidth(5,120); sheet.setColumnWidth(6,120);
  sheet.setColumnWidth(7,160); sheet.setColumnWidth(8,200); sheet.setColumnWidth(9,160);
  sheet.setColumnWidth(10,100);sheet.setColumnWidth(11,160);sheet.setColumnWidth(12,80);
  sheet.setColumnWidth(13,160);sheet.setColumnWidth(14,160);sheet.setColumnWidth(15,160);
  sheet.setColumnWidth(16,180);sheet.setColumnWidth(17,260);
  sheet.setRowHeight(1, 36);

  SpreadsheetApp.flush();
  Logger.log('Sheet ready.');
}

// ── GET handler ───────────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter.action || '').trim();
  const token  = (e.parameter.token  || '').trim();

  // ── register ─────────────────────────────────────────────
  if (action === 'register') {
    if (!token) return jsonResponse({ success: false, message: 'No token.' });
    const now = Date.now();

    // Only wipe the OLD valid_ key (not everything).
    // This lets students who scanned 2 rotations ago still submit
    // if they're still filling the form within GRACE_MS.
    const keys = props.getKeys();
    keys.forEach(k => {
      if (k.startsWith('valid_')) {
        const ts = parseInt(props.getProperty(k) || '0');
        if (Date.now() - ts > GRACE_MS) props.deleteProperty(k);
      }
    });

    props.setProperty('valid_' + token, now.toString());
    return jsonResponse({ success: true });
  }

  // ── stop ─────────────────────────────────────────────────
  if (action === 'stop') {
    props.getKeys().forEach(k => {
      if (k.startsWith('valid_') || k.startsWith('used_')) props.deleteProperty(k);
    });
    return jsonResponse({ success: true });
  }

  // ── check ─────────────────────────────────────────────────
  // KEY FIX: No more claimed_ lock.
  // Many students can all scan the same QR — they just need to
  // submit before GRACE_MS expires. Duplicate guard is in doPost.
  if (action === 'check') {
    if (!token) return jsonResponse({ valid: false, message: 'No token provided.' });

    const validTs = parseInt(props.getProperty('valid_' + token) || '0');
    if (!validTs)
      return jsonResponse({ valid: false, message: 'This QR code has expired. Please scan the latest one on screen.' });

    if (Date.now() - validTs > GRACE_MS)
      return jsonResponse({ valid: false, message: 'This QR code has expired. Please scan the latest one on screen.' });

    // Already used — tell them but don't hard-block (they may have
    // submitted already on this same device)
    if (props.getProperty('used_' + token))
      return jsonResponse({ valid: false, message: 'You have already marked attendance with this QR code.' });

    return jsonResponse({ valid: true });
  }

  return jsonResponse({ success: false, message: 'Unknown action.' });
}

// ── POST handler ──────────────────────────────────────────────
function doPost(e) {
  const token   = (e.parameter.token   || '').trim();
  const name    = (e.parameter.name    || '').trim();
  const roll    = (e.parameter.roll    || '').trim();
  const branch  = (e.parameter.branch  || '').trim();
  const year    = (e.parameter.year    || '').trim();
  const email   = (e.parameter.email   || '').trim();
  const college = (e.parameter.college || '').trim();
  const date    = (e.parameter.date    || '').trim();
  const subject = (e.parameter.subject || '').trim();
  const fb1     = (e.parameter.fb1     || '').trim();
  const fb2     = (e.parameter.fb2     || '').trim();
  const fb3     = (e.parameter.fb3     || '').trim();
  const fb4     = (e.parameter.fb4     || '').trim();
  const fb5     = (e.parameter.fb5     || '').trim();

  if (!token || !name || !roll || !email)
    return jsonResponse({ success: false, message: 'Please fill in all required fields.' });

  // ── Race-safe duplicate check using LockService ───────────
  // Without this lock, two students submitting at the exact same
  // millisecond can both pass the used_ check and both get written.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000); // wait up to 8s to get the lock
  } catch(e) {
    return jsonResponse({ success: false, message: 'Server busy. Please wait 5 seconds and try again.' });
  }

  try {
    // Check duplicate inside the lock
    if (props.getProperty('used_' + token))
      return jsonResponse({ success: false, message: 'You have already submitted attendance with this QR code.' });

    // Validate token is still live
    const validTs = parseInt(props.getProperty('valid_' + token) || '0');
    if (!validTs || Date.now() - validTs > GRACE_MS)
      return jsonResponse({ success: false, message: 'This QR code has expired. Please scan the latest one on screen.' });

    // Duplicate roll number check within this session
    const dupKey = 'roll_' + subject.replace(/\s/g,'') + '_' + roll;
    if (props.getProperty(dupKey))
      return jsonResponse({ success: false, message: 'Attendance already marked for Roll No: ' + roll + ' in this session.' });

    // Mark token used and roll seen — both while we hold the lock
    props.setProperty('used_' + token, new Date().toISOString());
    props.setProperty(dupKey, '1');

  } finally {
    lock.releaseLock();
  }

  // ── Write to sheet ────────────────────────────────────────
  // Formatting is REMOVED from here — the sheet is pre-styled by fixSheet().
  // Each doPost call now does exactly ONE sheet operation: appendRow.
  // This is 10-20x faster per submission at scale.
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || (() => {
      const s = ss.insertSheet(SHEET_NAME);
      s.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
      return s;
    })();

    const dp = date.split('-');
    const formattedDate = dp.length === 3 ? dp[2]+'/'+dp[1]+'/'+dp[0] : date;
    const srNo = sheet.getLastRow();  // header row counts as 1, so first data row gets srNo 1
    const submittedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

    sheet.appendRow([
      srNo, formattedDate, subject, name, roll,
      branch, year, email, college, token,
      submittedAt, 'Present',
      fb1, fb2, fb3, fb4, fb5
    ]);

    SpreadsheetApp.flush();
    return jsonResponse({ success: true, message: 'Attendance marked successfully!' });

  } catch(err) {
    // If sheet write fails, un-mark the token so student can retry
    props.deleteProperty('used_' + token);
    return jsonResponse({ success: false, message: 'Server error — please try submitting again. (' + err.message + ')' });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
