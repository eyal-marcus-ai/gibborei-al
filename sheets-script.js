// ══ Google Apps Script — גיבורי על: בקשות פרקים ══
// הוראות:
// 1. פתחו Google Sheets חדש, קראו לו "גיבורי על — בקשות פרקים"
// 2. Extensions > Apps Script
// 3. מחקו את כל הקוד הקיים, הדביקו את הקוד הזה
// 4. לחצו Save, ואז Deploy > New deployment
// 5. Type: Web app | Execute as: Me | Who has access: Anyone
// 6. לחצו Deploy, העתיקו את ה-URL שמתקבל
// 7. שלחו את ה-URL לאייל (או לקלוד) לעדכון הטופס

const SHEET_NAME = "בקשות";

function doPost(e) {
  const sheet = getOrCreateSheet();
  const data = e.parameter;
  const row = [
    new Date().toLocaleString("he-IL"),
    data.character   || "",
    data.child_name  || "",
    data.email       || "",
  ];
  sheet.appendRow(row);
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["תאריך", "גיבור מבוקש", "שם הילד", "מייל"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  }
  return sheet;
}
