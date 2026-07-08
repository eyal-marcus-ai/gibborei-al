// ══ Google Apps Script — גיבורי על: טופס אחד מאוחד מהאתר ══
// מה זה עושה (מודל מאוחד, 08/07/2026):
//   כל מי שממלא את הטופס (מייל + בקשה/מילה טובה) →
//     1. נכנס לרשימת העדכונים בסמוב (Superheros, 1145138)
//     2. נרשם בטאב "נרשמים לעדכונים" (עם הבקשה)
//     3. נשלח מייל התראה לאייל (עם הבקשה)
//   (type:'signup' ו-type:'request' נשמרים לתאימות לאחור עם גרסאות אתר ישנות)
// התקנה: הגיליון > Extensions > Apps Script > הדבק > Deploy (New version) לאותו /exec.

const REQUESTS_SHEET_GID = 1783764722; // טאב בקשות הפרקים (legacy)
const SIGNUPS_SHEET_NAME = "נרשמים לעדכונים";
const NOTIFY_EMAIL = "eyal@eyalmarcus.com";
const SMOOVE_KEY = "30018ed0-8f29-49a1-9d5f-160efeafb499";
const SMOOVE_LIST_ID = 1145138; // רשימת Superheros בסמוב

function doPost(e) {
  let data = {};
  try { data = JSON.parse(e.postData.contents); } catch (err) { data = (e && e.parameter) || {}; }
  try {
    if (data.type === "subscribe" || data.type === "signup") handleSubscribe(data);
    else if (data.type === "request") handleRequest(data); // legacy
  } catch (err) {
    MailApp.sendEmail(NOTIFY_EMAIL, "שגיאה בטופס באתר גיבורי על",
      "משהו נכשל בעיבוד טופס מהאתר.\n\nשגיאה: " + err + "\n\nהנתונים שהגיעו:\n" + JSON.stringify(data, null, 2));
  }
  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── הטופס המאוחד: מייל + בקשה/מילה טובה → סמוב + גיליון נרשמים + מייל ──
function handleSubscribe(data) {
  const email = data.email || "";
  const message = data.message || data.request || "";
  const name = data.name || ""; // בדרך כלל ריק בטופס המאוחד
  const status = subscribeToUpdates(name, email, message);

  MailApp.sendEmail(NOTIFY_EMAIL,
    "נרשם/ה חדש/ה לגיבורי על - " + (email || "בלי מייל"),
    "מישהו מילא את הטופס באתר גיבורי על!\n\n" +
    "מייל: " + email + "\n" +
    (message ? "הבקשה / ההודעה:\n" + message + "\n\n" : "\n") +
    "נכנס לרשימת העדכונים בסמוב? " + status + "\n\n" +
    "טאב הנרשמים:\n" +
    "https://docs.google.com/spreadsheets/d/16Sv2Y5FGL_2kvtlQN7_iFhhwha0-B1Q5YRKi08gDUcI/edit");
}

// ── הוספה לרשימת העדכונים: סמוב 1145138 + שורה בטאב נרשמים (כולל הבקשה). מחזיר סטטוס סמוב ──
function subscribeToUpdates(name, email, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SIGNUPS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SIGNUPS_SHEET_NAME);
    sheet.appendRow(["תאריך", "שם", "מייל", "נכנס לסמוב?", "בקשה / הודעה"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  } else if (!sheet.getRange(1, 5).getValue()) {
    sheet.getRange(1, 5).setValue("בקשה / הודעה").setFontWeight("bold"); // הוסף כותרת עמודה חדשה פעם אחת
  }

  let smooveStatus = "כן";
  try {
    const res = UrlFetchApp.fetch("https://rest.smoove.io/v1/Contacts?listId=" + SMOOVE_LIST_ID + "&updateIfExists=true", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + SMOOVE_KEY },
      payload: JSON.stringify({
        email: email,
        firstName: name || "",
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

  sheet.appendRow([new Date().toLocaleString("he-IL", { timeZone: "Asia/Jerusalem" }), name || "", email || "", smooveStatus, message || ""]);
  return smooveStatus;
}

// ── legacy: בקשת פרק מובנית (character/child_name) — נשמר לתאימות עם גרסת אתר ישנה ──
function handleRequest(data) {
  const sheet = getSheetByGid(REQUESTS_SHEET_GID);
  const lastRow = sheet.getLastRow();
  const lastNum = Number(sheet.getRange(lastRow, 1).getValue()) || 0;
  const num = lastNum + 1;
  sheet.appendRow([num, data.child_name || "", data.character || "", data.email || ""]);
  let updatesLine = "";
  if (data.wants_updates && data.email) {
    const status = subscribeToUpdates(data.child_name, data.email, "בקשת פרק: " + (data.character || ""));
    updatesLine = "\nסימן/ה לקבל עדכונים → נכנס לסמוב? " + status + "\n";
  }
  MailApp.sendEmail(NOTIFY_EMAIL,
    "בקשת פרק חדשה (#" + num + ") - " + (data.character || "בלי שם"),
    "בקשה חדשה מהאתר גיבורי על!\n\n" +
    "גיבור/נבל: " + (data.character || "") + "\n" +
    "שם הילד: " + (data.child_name || "") + "\n" +
    "מייל: " + (data.email || "") + "\n" +
    updatesLine + "\n" +
    "נוספה לגיליון כשורה מספר " + num + ":\n" +
    "https://docs.google.com/spreadsheets/d/16Sv2Y5FGL_2kvtlQN7_iFhhwha0-B1Q5YRKi08gDUcI/edit#gid=" + REQUESTS_SHEET_GID);
}

function getSheetByGid(gid) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(function (s) { return s.getSheetId() === gid; });
  if (!sheet) throw new Error("לא נמצא טאב עם gid " + gid);
  return sheet;
}
