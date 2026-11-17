/**
 * OPTIONAL: collects submissions from the OASIS Staff Details page in a Google Sheet.
 * The sheet "OASIS Staff Details (Responses)" is created automatically in your
 * Google Drive when the first teacher submits.
 *
 * 1. Open script.google.com -> New project -> paste this file -> Save.
 * 2. Deploy -> New deployment -> Web app -> Execute as: Me, Who has access: Anyone.
 * 3. Copy the web-app URL into SHEET_URL near the bottom of index.html.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('SHEET_ID');
    var book;
    if (id) {
      book = SpreadsheetApp.openById(id);
    } else {
      book = SpreadsheetApp.create('OASIS Staff Details (Responses)');
      props.setProperty('SHEET_ID', book.getId());
    }
    var sheet = book.getSheets()[0];
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
