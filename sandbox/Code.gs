/**
 * Apps Script מאוחד — "מִשְמָר" סנדבוקס
 * ------------------------------------------------------------
 * גרסה 2: תומך בכמה סוגי טפסים באותו endpoint אחד.
 * כל רשומה חייבת לכלול שדה formType (למשל "kabala_homre_gelem", "gmp_checklist"
 * וכו') — כל סוג טופס נכתב ללשונית (tab) נפרדת בגיליון, אבל כולם משתמשים
 * באותו WEBHOOK_URL. כך אפשר לחבר טפסים נוספים בעתיד בלי להקים
 * Apps Script/גיליון חדשים בכל פעם.
 *
 * doPost  — מקבל רשומה בודדת (עם formType) ומוסיף אותה ללשונית המתאימה
 * doGet   — מחזיר רשומות לפי שאילתה: formType, from/to, status, ועוד,
 *           עם ?formType=... חובה כדי לדעת מאיזו לשונית לקרוא
 *
 * התקנה / עדכון:
 * 1. בגיליון הקיים (הסנדבוקס) → Extensions → Apps Script
 * 2. מחק את כל התוכן הקיים, הדבק את כל הקובץ הזה, שמור
 * 3. Deploy → Manage deployments → עיפרון עריכה → Version: New version → Deploy
 *    (ה-URL נשאר זהה — אין צורך לעדכן שום טופס קיים)
 */

// עמודות משותפות לכל סוגי הטפסים. כל שדה נוסף וספציפי לטופס נשמר
// כ-JSON בעמודה extra, כדי לא להגביל מבנה לכל טופס עתידי.
const COMMON_HEADERS = ["timestamp", "formType", "date", "time", "status", "extra"];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const formType = data.formType || "unknown";
    const sheet = getOrCreateSheet(formType);

    const extra = {};
    Object.keys(data).forEach(function (k) {
      if (COMMON_HEADERS.indexOf(k) === -1 && k !== "formType") extra[k] = data[k];
    });

    const row = [
      new Date().toISOString(),      // timestamp — נקבע תמיד בשרת
      formType,
      data.date || "",
      data.time || "",
      data.status || "",
      JSON.stringify(extra),
    ];

    sheet.appendRow(row);
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    const formType = p.formType;
    if (!formType) {
      return jsonResponse({ ok: false, error: "חובה לציין formType" });
    }

    const sheet = getOrCreateSheet(formType);
    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      return jsonResponse({ ok: true, count: 0, records: [] });
    }

    const headers = values[0];
    let rows = values.slice(1).map(function (r) {
      const obj = {};
      headers.forEach(function (h, i) {
        let v = r[i];
        if (v instanceof Date) {
          if (h === "time") {
            v = Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");
          } else if (h === "timestamp") {
            v = v.toISOString();
          } else {
            v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
        }
        obj[h] = v;
      });
      // פותח את שדות ה-extra בחזרה לשדות רגילים ברמה העליונה
      if (obj.extra) {
        try {
          const parsed = JSON.parse(obj.extra);
          Object.keys(parsed).forEach(function (k) { obj[k] = parsed[k]; });
        } catch (e2) {}
        delete obj.extra;
      }
      return obj;
    });

    if (p.from) rows = rows.filter(function (r) { return String(r.date) >= p.from; });
    if (p.to) rows = rows.filter(function (r) { return String(r.date) <= p.to; });
    if (p.status) rows = rows.filter(function (r) { return r.status === p.status; });
    if (p.material) rows = rows.filter(function (r) { return r.material === p.material; });
    if (p.supplier) rows = rows.filter(function (r) { return r.supplier === p.supplier; });
    if (p.inspector) rows = rows.filter(function (r) { return r.inspector === p.inspector; });
    if (p.limit) {
      const n = Number(p.limit);
      rows = rows.slice(Math.max(0, rows.length - n));
    }

    return jsonResponse({ ok: true, count: rows.length, records: rows });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function getOrCreateSheet(formType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabName = String(formType).substring(0, 90); // הגבלת אורך שם לשונית
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(COMMON_HEADERS);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
