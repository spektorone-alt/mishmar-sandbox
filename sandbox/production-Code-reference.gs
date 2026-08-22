/**
 * Code.gs — בק-אנד משותף לכל 8 הטפסים של דו"ח בקרת תהליך ייצור בצקים קפואים (טופס 1023)
 * חובה: להריץ מתוך Extensions > Apps Script בתוך ה-Google Sheet עצמו (container-bound),
 * ואז Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * לעדכונים עתידיים: Manage deployments > ערוך > New version (לא ליצור Deployment חדש!).
 *
 * ------------------------------------------------------------------
 * עדכון: הוספת תמיכה בפרמטר ?date=YYYY-MM-DD אופציונלי ב-doGet, כדי לאפשר
 * צפייה בדוחות של ימים קודמים (לא רק היום). כשלא מציינים date, ההתנהגות
 * זהה לחלוטין למה שהיה קודם - שום דבר קיים לא השתנה.
 * שתי הפונקציות find DailyFile_ ו-parseDateParam_ הן תוספת בלבד.
 * ------------------------------------------------------------------
 */

// שם הטאב בגיליון עבור כל formType
const SHEET_MAP = {
  plan:      '1. תוכנית עבודה יומית',
  weighing:  '2. בקרת שקילה והכנת מתכון',
  kneading:  '3. בקרת לישה',
  rounding:  '4. בקרת כדרור',
  proofing:  '5. בקרת התפחה',
  baking:    '6. בקרת אפייה חלקית',
  packaging: '8. בקרת אריזה וסימון',
  finalcheck:'8א. בדיקת מוצר סופי לאחר אריזה'
};

// תיקיית Drive לקבצי האקסל היומיים של טופס 2 (בקרת שקילה והכנת מתכון)
const KIT_DRIVE_FOLDER_ID = '1JOtQII-vLygx8XE_vZsszIgrCaBASYD4';
// תיקיית Drive לקבצי האקסל היומיים של טופס 3 (בקרת לישה)
const LISHA_DRIVE_FOLDER_ID = '1iPtfOiWC6U-6i0A8-1rfAcs-Txb_iEOO';
// תיקיית Drive לקבצי האקסל היומיים של טופס 1 (תוכנית עבודה יומית)
const PLAN_DRIVE_FOLDER_ID = '1_KN-Lw-G7Z6I3KtyaJ40WQuH19pC-V4K';
// תיקיית Drive לקבצי האקסל היומיים של טופס 8 (בקרת אריזה וסימון)
const PACKAGING_DRIVE_FOLDER_ID = '1rcu6jB_Vu5VnAX6WUiT6x8wUjv0NuAEi';
const HEBREW_DAY_NAMES = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];

// כותרות עמודות לכל טאב (בסדר הכתיבה בפועל)
const HEADERS_MAP = {
  plan: ["מס'", 'מוצר', 'סוג מוצר (כדור,פוקציה..)', 'משקל יחידה מתוכנן',
         'כמות מתוכננת (כדורים)', 'אפייה חלקית (כן/לא)', 'סוג אריזה', 'הערות',
         'נרשם ע"י', 'זמן רישום'],
  weighing: ['מתכון', 'חו"ג', "מס' אצווה של חו\"ג", "מס' אצווה של חומר הגלם עצמו", 'תוקף חו"ג', 'הוכנס למיקסר',
             'כמות קיטים סופי', 'חתימה'],
  packaging: ["מס' עגלה", 'כמות יחידות באריזה', 'סוג הבצק', 'משקל יחידה', 'אריזה בודדת', 'סוג האריזה',
              'צבע האריזה', 'צבע סרט סגירה', 'כמות אריזות סופית', 'תאריך ייצור', 'תאריך תפוגה',
              'סידורי האריזה', 'שם מאשר', 'תאריך רשימה'],
  kneading: ['מספר לישה', 'מוצר', "מס' אצווה מנה", 'מספר אצווה קיט', 'זמן לישה', "טמפ' יעד",
             "טמפ' בפועל", 'משקל בצק', 'תקין/לא תקין', 'שעה הפרשת חלה', 'חתימה']
};

// מוצא (או יוצר) קובץ "אקסל" (Google Sheet) של יום הייצור הנוכחי בתיקיית Drive נתונה
function getOrCreateDailyFile_(folderId, namePrefix, sheetTabName, headers) {
  const folder = DriveApp.getFolderById(folderId);
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const dayName = HEBREW_DAY_NAMES[now.getDay()];
  const dateStr = Utilities.formatDate(now, tz, 'dd-MM-yyyy');
  const fileName = namePrefix + ' ' + dayName + ' ' + dateStr;

  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) {
    const file = existing.next();
    return SpreadsheetApp.openById(file.getId());
  }

  const ss = SpreadsheetApp.create(fileName);
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  const rootParents = file.getParents();
  while (rootParents.hasNext()) {
    const p = rootParents.next();
    if (p.getId() !== folder.getId()) p.removeFile(file);
  }

  const sheet = ss.getActiveSheet();
  sheet.setName(sheetTabName);
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return ss;
}

// === תוספת חדשה: מוצא קובץ של תאריך נתון (לא היום), בלי ליצור אם לא קיים ===
// מחזיר null אם אין קובץ לאותו יום (כלומר לא היה ייצור/לא מולא טופס אז)
function findDailyFile_(folderId, namePrefix, dateObj) {
  const folder = DriveApp.getFolderById(folderId);
  const tz = Session.getScriptTimeZone();
  const dayName = HEBREW_DAY_NAMES[dateObj.getDay()];
  const dateStr = Utilities.formatDate(dateObj, tz, 'dd-MM-yyyy');
  const fileName = namePrefix + ' ' + dayName + ' ' + dateStr;

  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) {
    return SpreadsheetApp.openById(existing.next().getId());
  }
  return null;
}

// === תוספת חדשה: הופך מחרוזת 'YYYY-MM-DD' לאובייקט Date מקומי (לא UTC) ===
function parseDateParam_(dateStr) {
  const parts = dateStr.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function getOrCreateKitFile_() {
  return getOrCreateDailyFile_(KIT_DRIVE_FOLDER_ID, 'קיטים', 'קיטים', HEADERS_MAP.weighing);
}

function getOrCreateLishaFile_() {
  return getOrCreateDailyFile_(LISHA_DRIVE_FOLDER_ID, 'לישה', 'לישה', HEADERS_MAP.kneading);
}

function getOrCreatePlanFile_() {
  return getOrCreateDailyFile_(PLAN_DRIVE_FOLDER_ID, 'תוכנית עבודה', 'תוכנית עבודה', HEADERS_MAP.plan);
}

function getOrCreatePackagingFile_() {
  return getOrCreateDailyFile_(PACKAGING_DRIVE_FOLDER_ID, 'אריזה', 'אריזה', HEADERS_MAP.packaging);
}

// פונקציית בדיקה ידנית - להריץ פעם אחת מתפריט ה-Run כדי לוודא שהגישה ל-Drive עובדת
function testKitFile() {
  const ss = getOrCreateKitFile_();
  Logger.log('נוצר/נפתח בהצלחה: ' + ss.getUrl());
}

// מונה גלובלי קבוע למס' אצווה מנה (לישה) - לעולם לא חוזר על עצמו, גם בין ימים/קבצים שונים
function getNextKneadBatchNum_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const current = parseInt(props.getProperty('kneadBatchCounter') || '5000000', 10);
    const next = current + 1;
    props.setProperty('kneadBatchCounter', String(next));
    return next;
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.parameter.payload);
    const formType = e.parameter.formType;

    // טופס 2 (בקרת שקילה והכנת מתכון) נשמר כקובץ אקסל נפרד לכל יום ייצור ב-Drive
    if (formType === 'weighing') {
      const ss = getOrCreateKitFile_();
      const sheet = ss.getSheetByName('קיטים');
      const headers = HEADERS_MAP.weighing;
      const row = headers.map(h => payload[h] !== undefined ? payload[h] : '');
      sheet.appendRow(row);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', file: ss.getName() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // טופס 3 (בקרת לישה) נשמר כקובץ אקסל נפרד לכל יום ייצור ב-Drive
    if (formType === 'kneading') {
      const ss = getOrCreateLishaFile_();
      const sheet = ss.getSheetByName('לישה');
      const headers = HEADERS_MAP.kneading;
      const row = headers.map(h => payload[h] !== undefined ? payload[h] : '');
      sheet.appendRow(row);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', file: ss.getName() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // טופס 1 (תוכנית עבודה יומית) נשמר כקובץ אקסל נפרד לכל יום ייצור ב-Drive
    if (formType === 'plan') {
      const ss = getOrCreatePlanFile_();
      const sheet = ss.getSheetByName('תוכנית עבודה');
      const headers = HEADERS_MAP.plan;
      const row = headers.map(h => payload[h] !== undefined ? payload[h] : '');
      sheet.appendRow(row);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', file: ss.getName() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // טופס 8 (בקרת אריזה וסימון) נשמר כקובץ אקסל נפרד לכל יום ייצור ב-Drive
    if (formType === 'packaging') {
      const ss = getOrCreatePackagingFile_();
      const sheet = ss.getSheetByName('אריזה');
      const headers = HEADERS_MAP.packaging;
      const row = headers.map(h => payload[h] !== undefined ? payload[h] : '');
      sheet.appendRow(row);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', file: ss.getName() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheetName = SHEET_MAP[formType];
    if (!sheetName) throw new Error('formType לא מוכר: ' + formType);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = HEADERS_MAP[formType];
      if (headers) sheet.appendRow(headers);
    }

    const headers = HEADERS_MAP[formType] || Object.keys(payload);
    const row = headers.map(h => payload[h] !== undefined ? payload[h] : '');
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET — משמש לטפסים הבאים בשרשרת כדי למשוך רשימת אצוות פתוחות מהיום (לתפריט בחירה)
// וגם לדשבורד לקריאה בלבד. תמיכה חדשה: ?date=YYYY-MM-DD לצפייה בימים קודמים.
function doGet(e) {
  try {
    const formType = e.parameter.formType || 'plan';
    const date = e.parameter.date; // 'YYYY-MM-DD'
    const action = e.parameter.action;

    // מחזיר את מספר האצווה (מנה) הבא - גלובלי, לעולם לא חוזר על עצמו
    if (formType === 'kneading' && action === 'nextBatch') {
      const next = getNextKneadBatchNum_();
      return ContentService.createTextOutput(JSON.stringify({ batchNum: next }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (formType === 'kneading') {
      // תוספת: אם ניתן date, מחפשים את קובץ אותו יום בלבד (בלי ליצור); אחרת - התנהגות ישנה (היום)
      const ss = date ? findDailyFile_(LISHA_DRIVE_FOLDER_ID, 'לישה', parseDateParam_(date)) : getOrCreateLishaFile_();
      if (!ss) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
      const sheet = ss.getSheetByName('לישה');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1).map(r => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = r[i]);
        return obj;
      });
      return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
    }

    if (formType === 'plan') {
      const ss = date ? findDailyFile_(PLAN_DRIVE_FOLDER_ID, 'תוכנית עבודה', parseDateParam_(date)) : getOrCreatePlanFile_();
      if (!ss) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
      const sheet = ss.getSheetByName('תוכנית עבודה');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const batchCol = headers.indexOf("מס'");
      const compactDate = date ? date.replace(/-/g, '') : null;
      const rows = data.slice(1)
        .filter(r => !compactDate || String(r[batchCol]).startsWith(compactDate))
        .map(r => {
          const obj = {};
          headers.forEach((h, i) => obj[h] = r[i]);
          return obj;
        });
      return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
    }

    if (formType === 'weighing') {
      const ss = date ? findDailyFile_(KIT_DRIVE_FOLDER_ID, 'קיטים', parseDateParam_(date)) : getOrCreateKitFile_();
      if (!ss) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
      const sheet = ss.getSheetByName('קיטים');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1).map(r => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = r[i]);
        return obj;
      });
      return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
    }

    if (formType === 'packaging') {
      const ss = date ? findDailyFile_(PACKAGING_DRIVE_FOLDER_ID, 'אריזה', parseDateParam_(date)) : getOrCreatePackagingFile_();
      if (!ss) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
      const sheet = ss.getSheetByName('אריזה');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1).map(r => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = r[i]);
        return obj;
      });
      return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
    }

    const sheetName = SHEET_MAP[formType];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const batchCol = headers.indexOf("מס'");
    const compactDate = date ? date.replace(/-/g, '') : null;
    const rows = data.slice(1)
      .filter(r => !compactDate || String(r[batchCol]).startsWith(compactDate))
      .map(r => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = r[i]);
        return obj;
      });

    return ContentService.createTextOutput(JSON.stringify(rows))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
