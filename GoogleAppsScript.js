// Code.gs - Google Apps Script Backend Engine (Seat Number & Unique IP Support)
// Copy & Paste this entire file into your Apps Script editor (Extensions -> Apps Script)

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Lock for up to 10 seconds to handle concurrent student submissions

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    var rollNo = "N/A";
    var name = "N/A";
    var seatNo = "N/A";
    var timestamp = new Date().toLocaleString();
    var ip = "Unknown";

    // 1. Read Form Parameter Submissions (x-www-form-urlencoded)
    if (e && e.parameter) {
      rollNo = e.parameter.rollNo || rollNo;
      name = e.parameter.name || name;
      seatNo = e.parameter.seatNo || seatNo;
      timestamp = e.parameter.timestamp || timestamp;
      ip = e.parameter.ip || ip;
    }

    // 2. Read JSON Submissions as fallback
    if ((rollNo === "N/A" || rollNo === "") && e && e.postData && e.postData.contents) {
      try {
        var data = JSON.parse(e.postData.contents);
        rollNo = data.rollNo || rollNo;
        name = data.name || name;
        seatNo = data.seatNo || seatNo;
        timestamp = data.timestamp || timestamp;
        ip = data.ip || ip;
      } catch (jsonErr) {}
    }

    // 3. Auto-create headers on Row 1 if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Roll Number", "Student Name", "Seat Number", "Status", "IP Address"]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    }

    // 4. Append student submission row into Google Sheet
    sheet.appendRow([timestamp, rollNo, name, seatNo, "Present", ip]);

    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.toString()).setMimeType(ContentService.MimeType.TEXT);
  } finally {
    lock.releaseLock();
  }
}

// Authorization Helper Function
// Run this function ONCE inside the Apps Script Editor (Select 'setupSheet' -> Click 'Run')
// to authorize Google Sheets permissions properly.
function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Roll Number", "Student Name", "Seat Number", "Status", "IP Address"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  }
}