/**
 * OPTIONAL: collects submissions from the OASIS Staff Details page in a Google Sheet.
 *
 * 1. Create a new Google Sheet -> Extensions -> Apps Script -> paste this file.
 * 2. Deploy -> New deployment -> Web app -> Execute as: Me, Who has access: Anyone.
 * 3. Copy the web-app URL into SHEET_URL near the bottom of index.html.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var keys = Object.keys(data);
    if (sheet.getLastRow() === 0) sheet.appendRow(['Submitted at'].concat(keys));
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = headers.map(function (h, i) {
      if (i === 0) return new Date();
      var v = data[h] == null ? '' : String(data[h]);
      return /^[=+\-@]/.test(v) ? "'" + v : v;   // keep text as text
    });
    sheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
