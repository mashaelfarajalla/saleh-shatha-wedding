/**
 * Backend for the wishes section in invitation.html.
 * Paste into Extensions → Apps Script of a new Google Sheet, then
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone),
 * and put the web-app URL in CONFIG.sheetsUrl.
 *
 * Columns: time | name | message | show.
 * Untick "show" to hide a wish from the page; delete the row to remove it.
 *
 * WhatsApp forwarding (optional): each new wish is sent to WA_PHONE through
 * CallMeBot. Get the key by sending "I allow callmebot to send me messages"
 * from that phone to the CallMeBot number listed on callmebot.com, then paste it in WA_APIKEY.
 */

const WA_PHONE = '+972592459391';
const WA_APIKEY = '';

function notifyWhatsApp_(name, message) {
  if (!WA_APIKEY) return;
  const text = '💌 تهنئة جديدة\n' + name + ':\n' + message;
  try {
    UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(WA_PHONE) +
      '&text=' + encodeURIComponent(text) + '&apikey=' + encodeURIComponent(WA_APIKEY), { muteHttpExceptions: true });
  } catch (err) {}
}

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(['time', 'name', 'message', 'show']);
  // Older sheets without the show column: add it, ticked for existing wishes.
  if (!sh.getRange(1, 4).getValue()) {
    sh.getRange(1, 4).setValue('show');
    const n = sh.getLastRow() - 1;
    if (n > 0) sh.getRange(2, 4, n).insertCheckboxes().setValue(true);
  }
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
    .filter(r => r[1] && r[2] && r[3] !== false)
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
    const sh = sheet_();
    sh.appendRow([new Date(), name, message]);
    sh.getRange(sh.getLastRow(), 4).insertCheckboxes().setValue(true);
    lock.releaseLock();
    notifyWhatsApp_(name.replace(/^'/, ''), message.replace(/^'/, ''));
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  }
}
