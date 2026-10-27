// Code.gs - Google Apps Script Backend Engine
// Copy & Paste this code into your Google Sheet's Apps Script Editor (Extensions -> Apps Script)

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Prevent concurrent writing conflicts

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Auto-create headers on first submission if empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Roll Number", "Student Name", "Status"]);
      sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    }

    // Append student submission row
    sheet.appendRow([
      data.timestamp,
      data.rollNo,
      data.name,
      "Present"
    ]);

    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err).setMimeType(ContentService.MimeType.TEXT);
  } finally {
    lock.releaseLock();
  }
}