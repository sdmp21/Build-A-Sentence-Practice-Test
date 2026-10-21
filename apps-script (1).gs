/**
 * Receives Build a Sentence submissions from index.html and appends one row
 * per student to the "Responses" tab of the Google Sheet this script is attached to.
 *
 * Setup: see README.md.
 */

const SHEET_NAME = 'Responses';

const HEADERS = [
  'Timestamp', 'Submission ID', 'Name', 'Email',
  'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10',
  'Seconds used', 'Timed out', 'Browser'
];

// Columns (0-based) stored as plain text so a typed "=" or "+" is never read as a formula.
const TEXT_COLUMNS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 16];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }

    // Ignore a submission that was already saved (the page may retry).
    const id = clean_(data.submissionId, 100);
    if (id && sheet.getLastRow() > 1) {
      const ids = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === id) return json_({ ok: true, duplicate: true });
      }
    }

    const answers = Array.isArray(data.answers) ? data.answers : [];
    const row = [
      new Date().toISOString(),
      id,
      clean_(data.name, 200),
      clean_(data.email, 200)
    ];
    for (let q = 0; q < 10; q++) row.push(clean_(answers[q], 500));
    row.push(Number(data.secondsUsed) || 0);
    row.push(data.timedOut ? 'Yes' : 'No');
    row.push(clean_(data.userAgent, 300));

    const target = sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length);
    target.setNumberFormats([row.map((_, i) => (TEXT_COLUMNS.indexOf(i) > -1 ? '@' : 'General'))]);
    target.setValues([row]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Opening the web app URL in a browser shows this, so you can confirm it is live.
function doGet() {
  return json_({ ok: true, status: 'ready' });
}

function clean_(value, max) {
  return String(value == null ? '' : value).slice(0, max);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
