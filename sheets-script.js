// ══ Google Apps Script — גיבורי על: שני הטפסים מהאתר ══
// מה זה עושה:
//   בקשת פרק  → שורה בטאב הבקשות הקיים + מייל התראה לאייל
//   הרשמה     → שורה בטאב "נרשמים לעדכונים" + הוספה לסמוב + מייל התראה לאייל
// איך מתקינים (חד-פעמי):
//   1. פותחים את הגיליון "רשימת בקשות פרקים"
//   2. Extensions > Apps Script
//   3. מדביקים את הקוד הזה במקום כל מה שיש
//   4. Deploy > New deployment > Web app | Execute as: Me | Who has access: Anyone
//   5. מעתיקים את ה-URL ומדביקים ב-index.html במשתנה GAS_URL

const REQUESTS_SHEET_GID = 1783764722; // הטאב הקיים של בקשות הפרקים
const SIGNUPS_SHEET_NAME = "נרשמים לעדכונים";
const NOTIFY_EMAIL = "eyal@eyalmarcus.com";
const SMOOVE_KEY = "30018ed0-8f29-49a1-9d5f-160efeafb499";
const SMOOVE_LIST_ID = 1145138; // רשימת Superheros בסמוב

function doPost(e) {
  let data = {};
  try { data = JSON.parse(e.postData.contents); } catch (err) { data = (e && e.parameter) || {}; }
  try {
    if (data.type === "signup") handleSignup(data);
    else if (data.type === "request") handleRequest(data);
  } catch (err) {
    MailApp.sendEmail(NOTIFY_EMAIL, "שגיאה בטופס באתר גיבורי על",
      "משהו נכשל בעיבוד טופס מהאתר.\n\nשגיאה: " + err + "\n\nהנתונים שהגיעו:\n" + JSON.stringify(data, null, 2));
  }
  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── בקשת פרק: נכנסת לטאב הקיים בפורמט הקיים (מס' | שם מלא | מה ביקש/ה | מייל) ──
function handleRequest(data) {
  const sheet = getSheetByGid(REQUESTS_SHEET_GID);
  const lastRow = sheet.getLastRow();
  const lastNum = Number(sheet.getRange(lastRow, 1).getValue()) || 0;
  const num = lastNum + 1;
  sheet.appendRow([num, data.child_name || "", data.character || "", data.email || ""]);
  MailApp.sendEmail(NOTIFY_EMAIL,
    "בקשת פרק חדשה (#" + num + ") - " + (data.character || "בלי שם"),
    "בקשה חדשה מהאתר גיבורי על!\n\n" +
    "גיבור/נבל: " + (data.character || "") + "\n" +
    "שם הילד: " + (data.child_name || "") + "\n" +
    "מייל: " + (data.email || "") + "\n\n" +
    "נוספה לגיליון כשורה מספר " + num + ":\n" +
    "https://docs.google.com/spreadsheets/d/16Sv2Y5FGL_2kvtlQN7_iFhhwha0-B1Q5YRKi08gDUcI/edit#gid=" + REQUESTS_SHEET_GID);
}

// ── הרשמה לעדכונים: טאב נרשמים + סמוב + מייל ──
function handleSignup(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SIGNUPS_SHEET_NAME);
    sheet.appendRow(["תאריך", "שם", "מייל", "נכנס לסמוב?"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  }

  let smooveStatus = "כן";
  try {
    const res = UrlFetchApp.fetch("https://rest.smoove.io/v1/Contacts?listId=" + SMOOVE_LIST_ID, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + SMOOVE_KEY },
      payload: JSON.stringify({
        email: data.email,
        firstName: data.name || "",
        lists_ToSubscribe: [SMOOVE_LIST_ID]
      }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() >= 300) {
      smooveStatus = "לא - שגיאה " + res.getResponseCode() + ": " + res.getContentText().slice(0, 200);
    }
  } catch (err) {
    smooveStatus = "לא - " + err;
  }

  sheet.appendRow([new Date().toLocaleString("he-IL", { timeZone: "Asia/Jerusalem" }), data.name || "", data.email || "", smooveStatus]);

  MailApp.sendEmail(NOTIFY_EMAIL,
    "נרשם חדש לעדכוני גיבורי על - " + (data.name || data.email || ""),
    "מישהו נרשם לעדכונים באתר!\n\n" +
    "שם: " + (data.name || "") + "\n" +
    "מייל: " + (data.email || "") + "\n" +
    "נכנס לסמוב? " + smooveStatus + "\n\n" +
    "הגיליון:\n" +
    "https://docs.google.com/spreadsheets/d/16Sv2Y5FGL_2kvtlQN7_iFhhwha0-B1Q5YRKi08gDUcI/edit");
}

function getSheetByGid(gid) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(function (s) { return s.getSheetId() === gid; });
  if (!sheet) throw new Error("לא נמצא טאב עם gid " + gid);
  return sheet;
}
