/**
 * Backend for the wishes section in invitation.html.
 * Paste into Extensions → Apps Script of a new Google Sheet, then
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone),
 * and put the web-app URL in CONFIG.sheetsUrl.
 *
 * Columns: time | name | message. Delete a row in the sheet to remove a wish.
 */

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(['time', 'name', 'message']);
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Stop guest text from being treated as a spreadsheet formula.
function clean_(s, max) {
  s = String(s || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const wishes = rows
    .filter(r => r[1] && r[2])
    .map(r => ({ time: r[0], name: String(r[1]).replace(/^'/, ''), message: String(r[2]).replace(/^'/, '') }))
    .reverse();
  return json_(wishes);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = clean_(d.name, 60);
    const message = clean_(d.message, 500);
    if (!name || !message) return json_({ ok: false });
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    sheet_().appendRow([new Date(), name, message]);
    lock.releaseLock();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  }
}
