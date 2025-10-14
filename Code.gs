/*
  Google Apps Script backend for x.plus.bd
  - Sheet name must be exactly: "Users Dashboard"
  - Exposes doPost(e) for analytics events (page_enter, page_exit)
  - Exposes doGet(e) with action=rows to read structured JSON

  Security note: This is designed for a public Web App deployment (execute as Me, accessible to Anyone with link)
*/

const SHEET_NAME = 'Users Dashboard';

function getSheet_() {
  const ss = SpreadsheetApp.openById('11trONTNEuQ4WfmzTynytufAsBhYzgCkCem3L5w2VUvo');
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  return sheet;
}

function getHeaders_() {
  return [
    'Timestamp','Unique User ID','First Visit DateTime','Last Visit DateTime','Last Exit DateTime','Last Visited Duration (minutes)','Total Visited Duration (minutes)','Total Visit Count','Device Type','Device Model / Name','Browser / WebView Info','Screen Resolution / Ratio','Timezone','IP Address','Country','City','Location (Latitude, Longitude)','Last Visited Page / Section Name','App Entry Time','App Exit Time','Live Status Icon','User Status (Text)','Days Since Last Visit','Account Age (Days)','Repeat Visit Type','Total Sessions','Active Period Range','Source (optional)','Notes / Admin Tag'
  ];
}

function ensureHeaderRow_(sheet) {
  const headers = getHeaders_();
  const rng = sheet.getRange(1, 1, 1, headers.length);
  const values = rng.getValues();
  const hasHeader = values[0].some(v => v);
  if (!hasHeader) {
    rng.setValues([headers]);
  } else {
    // Ensure width matches
    const existing = values[0];
    if (existing.length < headers.length) {
      sheet.insertColumnsAfter(existing.length, headers.length - existing.length);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
}

function findRowByUid_(sheet, uid) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const uidCol = 2; // B: Unique User ID
  const range = sheet.getRange(2, uidCol, lastRow - 2 + 1, 1);
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if ((values[i][0] || '').toString() === uid) return 2 + i; // row index
  }
  return null;
}

function emptyRow_(sheet) {
  const headers = getHeaders_();
  const arr = new Array(headers.length).fill('');
  return arr;
}

function setRowValues_(sheet, rowIndex, obj) {
  const headers = getHeaders_();
  const row = emptyRow_(sheet);
  for (let i = 0; i < headers.length; i++) {
    const key = headers[i];
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      row[i] = obj[key];
    }
  }
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function updateRowValues_(sheet, rowIndex, obj) {
  const headers = getHeaders_();
  const existing = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const key = headers[i];
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      existing[i] = obj[key];
    }
  }
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([existing]);
}

function upsertEnter_(payload) {
  const sheet = getSheet_();
  ensureHeaderRow_(sheet);
  const uid = (payload.uid || '').toString();
  if (!uid) throw new Error('uid missing');
  const nowStr = new Date();

  const rowIndex = findRowByUid_(sheet, uid);
  const base = {
    'Timestamp': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    'Unique User ID': uid,
    'First Visit DateTime': payload.firstVisitIso ? Utilities.formatDate(new Date(payload.firstVisitIso), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
    'Last Visit DateTime': payload.entryTime ? Utilities.formatDate(new Date(payload.entryTime), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
    'Device Type': payload.deviceType || '',
    'Device Model / Name': payload.deviceModel || '',
    'Browser / WebView Info': payload.userAgent || '',
    'Screen Resolution / Ratio': payload.screenResolution || '',
    'Timezone': payload.timezone || '',
    'IP Address': payload.ip || '',
    'Country': payload.country || '',
    'City': payload.city || '',
    'Location (Latitude, Longitude)': payload.location || '',
    'Last Visited Page / Section Name': payload.sectionName || '',
    'App Entry Time': payload.entryTime ? Utilities.formatDate(new Date(payload.entryTime), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
  };

  if (rowIndex) {
    // Update existing row
    const prev = sheet.getRange(rowIndex, 1, 1, getHeaders_().length).getValues()[0];
    const visitCount = Number(prev[7] || 0) + 1; // H: Total Visit Count
    const totalSessions = Number(prev[26] || 0) + 1; // AA is index 27 but 0-based 26
    base['Total Visit Count'] = visitCount;
    base['Total Sessions'] = totalSessions;
    updateRowValues_(sheet, rowIndex, base);
  } else {
    // Insert new row at end
    const lastRow = Math.max(2, sheet.getLastRow() + 1);
    const obj = Object.assign({}, base, {
      'First Visit DateTime': base['First Visit DateTime'] || base['Last Visit DateTime'],
      'Total Visit Count': 1,
      'Total Sessions': 1,
    });
    setRowValues_(sheet, lastRow, obj);
  }
}

function upsertExit_(payload) {
  const sheet = getSheet_();
  ensureHeaderRow_(sheet);
  const uid = (payload.uid || '').toString();
  if (!uid) throw new Error('uid missing');
  const rowIndex = findRowByUid_(sheet, uid);
  const exitIso = payload.exitTime;
  const dur = Number(payload.durationMinutes || 0);
  if (!rowIndex) {
    // If exit comes first unexpectedly, create minimal row
    const lastRow = Math.max(2, sheet.getLastRow() + 1);
    const obj = {
      'Timestamp': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      'Unique User ID': uid,
      'Last Exit DateTime': exitIso ? Utilities.formatDate(new Date(exitIso), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
      'Last Visited Duration (minutes)': dur,
      'Total Visited Duration (minutes)': dur,
      'App Exit Time': exitIso ? Utilities.formatDate(new Date(exitIso), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
    };
    setRowValues_(sheet, lastRow, obj);
    return;
  }

  const headers = getHeaders_();
  const rowVals = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const totalIdx = headers.indexOf('Total Visited Duration (minutes)');
  const totalPrev = Number(rowVals[totalIdx] || 0);
  const newTotal = Math.round((totalPrev + dur) * 100) / 100;

  updateRowValues_(sheet, rowIndex, {
    'Timestamp': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    'Last Exit DateTime': exitIso ? Utilities.formatDate(new Date(exitIso), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
    'Last Visited Duration (minutes)': dur,
    'Total Visited Duration (minutes)': newTotal,
    'App Exit Time': exitIso ? Utilities.formatDate(new Date(exitIso), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
  });
}

function doPost(e) {
  try {
    const body = (e && e.parameter && e.parameter.payload) ? e.parameter.payload : (e && e.postData && e.postData.contents);
    if (!body) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'missing body' })).setMimeType(ContentService.MimeType.JSON);
    const data = typeof body === 'string' ? JSON.parse(body) : body;
    const event = data.event;
    if (event === 'page_enter') {
      upsertEnter_(data);
      return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    }
    if (event === 'page_exit') {
      upsertExit_(data);
      return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unknown event' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === 'rows') {
      const sheet = getSheet_();
      ensureHeaderRow_(sheet);
      const values = sheet.getDataRange().getValues();
      if (values.length <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ rows: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      const headers = values[0];
      const rows = values.slice(1).filter(r => r.some(x => x !== '' && x != null)).map(cols => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = cols[i]);
        return obj;
      });
      return ContentService.createTextOutput(JSON.stringify({ rows })).setMimeType(ContentService.MimeType.JSON);
    }

    return HtmlService.createHtmlOutput('x.plus.bd Apps Script online');
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}
