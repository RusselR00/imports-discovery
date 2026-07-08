// Google Apps Script — receives form POSTs and appends to the active sheet.
// Deploy: Extensions ▸ Apps Script, paste this, Deploy ▸ New deployment ▸ Web app.

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Responses')
             || SpreadsheetApp.getActiveSpreadsheet().insertSheet('Responses');

    var data = JSON.parse(e.postData.contents);

    // Fixed column order so every row lines up.
    var cols = ['timestamp','Name','Company','Email','Phone'];
    for (var i = 1; i <= 20; i++) cols.push('Q' + i);

    // Write header row once.
    if (sheet.getLastRow() === 0) sheet.appendRow(cols);

    var row = cols.map(function (c) { return data[c] || ''; });
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
