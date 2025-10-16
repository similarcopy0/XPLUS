/**
 * Users Analytics Tracker + JSON API for "Users Dashboard" sheet
 * - Handles POST events: page_enter, page_exit
 * - Exposes GET ?format=json for dashboard
 * - Auto-creates sheet/header and sets ArrayFormulas for derived columns
 * - Enhanced with better error handling, logging, and performance optimizations
 *
 * Deploy as Web App (Anyone with link).
 */

const CONFIG = {
  DEFAULT_SPREADSHEET_ID: '11trONTNEuQ4WfmzTynytufAsBhYzgCkCem3L5w2VUvo',
  SHEET_NAME: 'Users Dashboard',
  MAX_RETRIES: 3,
  LOCK_TIMEOUT: 15000,
  CACHE_DURATION: 300000, // 5 minutes
  MAX_ROWS_EXPORT: 10000,
  LOG_LEVEL: 'INFO' // DEBUG, INFO, WARN, ERROR
};

// Column headers in exact order
const HEADERS = [
  'Timestamp',
  'Unique User ID',
  'First Visit DateTime',
  'Last Visit DateTime',
  'Last Exit DateTime',
  'Last Visited Duration (minutes)',
  'Total Visited Duration (minutes)',
  'Total Visit Count',
  'Device Type',
  'Device Model / Name',
  'Browser / WebView Info',
  'Screen Resolution / Ratio',
  'Timezone',
  'IP Address',
  'Country',
  'City',
  'Location (Latitude, Longitude)',
  'Last Visited Page / Section Name',
  'App Entry Time',
  'App Exit Time',
  'Live Status Icon',
  'User Status (Text)',
  'Days Since Last Visit',
  'Account Age (Days)',
  'Repeat Visit Type',
  'Total Sessions',
  'Active Period Range',
  'Source (optional)',
  'Notes / Admin Tag'
];

const COL = (() => {
  const map = {};
  HEADERS.forEach((h, i) => (map[h] = i + 1)); // 1-based
  return Object.freeze(map);
})();

// Cache for performance
const cache = {
  sheetData: null,
  lastFetch: 0,
  spreadsheet: null
};

/**
 * Enhanced logging system
 */
class Logger {
  static log(level, message, data = null) {
    const levels = { DEBUG: 0, INFO: 1, WARN: 2, ERROR: 3 };
    const configLevel = levels[CONFIG.LOG_LEVEL] || 1;
    const messageLevel = levels[level] || 1;
    
    if (messageLevel >= configLevel) {
      const timestamp = new Date().toISOString();
      const logMessage = `[${timestamp}] ${level}: ${message}`;
      
      if (data) {
        console.log(logMessage, data);
      } else {
        console.log(logMessage);
      }
      
      // Store critical errors for debugging
      if (level === 'ERROR') {
        try {
          const errorSheet = getOrCreateErrorLogSheet();
          errorSheet.appendRow([timestamp, level, message, JSON.stringify(data || {})]);
        } catch (e) {
          console.error('Failed to log error to sheet:', e);
        }
      }
    }
  }
  
  static debug(message, data) { this.log('DEBUG', message, data); }
  static info(message, data) { this.log('INFO', message, data); }
  static warn(message, data) { this.log('WARN', message, data); }
  static error(message, data) { this.log('ERROR', message, data); }
}

/**
 * Get or create error log sheet
 */
function getOrCreateErrorLogSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.DEFAULT_SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Error Log');
  if (!sheet) {
    sheet = ss.insertSheet('Error Log');
    sheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Level', 'Message', 'Data']]);
  }
  return sheet;
}

/**
 * Enhanced validation utilities
 */
class Validator {
  static isValidUID(uid) {
    return typeof uid === 'string' && uid.trim().length > 0 && uid.length <= 100;
  }
  
  static isValidDate(dateStr) {
    if (!dateStr) return false;
    const date = new Date(dateStr);
    return !isNaN(date.getTime()) && date.getTime() > 0;
  }
  
  static isValidNumber(value, min = 0, max = Number.MAX_SAFE_INTEGER) {
    const num = Number(value);
    return !isNaN(num) && num >= min && num <= max;
  }
  
  static sanitizeString(str, maxLength = 255) {
    if (typeof str !== 'string') return '';
    return str.trim().substring(0, maxLength);
  }
  
  static validatePageEnterData(body) {
    const errors = [];
    
    if (!this.isValidUID(body.uid)) {
      errors.push('Invalid or missing UID');
    }
    
    if (body.entryTime && !this.isValidDate(body.entryTime)) {
      errors.push('Invalid entryTime format');
    }
    
    if (body.firstVisitIso && !this.isValidDate(body.firstVisitIso)) {
      errors.push('Invalid firstVisitIso format');
    }
    
    if (body.durationMinutes && !this.isValidNumber(body.durationMinutes, 0, 10080)) { // max 1 week
      errors.push('Invalid durationMinutes value');
    }
    
    return errors;
  }
  
  static validatePageExitData(body) {
    const errors = [];
    
    if (!this.isValidUID(body.uid)) {
      errors.push('Invalid or missing UID');
    }
    
    if (body.exitTime && !this.isValidDate(body.exitTime)) {
      errors.push('Invalid exitTime format');
    }
    
    if (body.durationMinutes && !this.isValidNumber(body.durationMinutes, 0, 10080)) {
      errors.push('Invalid durationMinutes value');
    }
    
    return errors;
  }
}

/**
 * Enhanced data processing utilities
 */
class DataProcessor {
  static sanitizeRowData(data) {
    const sanitized = {};
    
    // Sanitize string fields
    const stringFields = ['uid', 'deviceType', 'deviceModel', 'userAgent', 'screenResolution', 
                         'timezone', 'ip', 'country', 'city', 'location', 'sectionName', 'source'];
    
    stringFields.forEach(field => {
      if (data[field]) {
        sanitized[field] = Validator.sanitizeString(data[field]);
      }
    });
    
    // Validate and sanitize numeric fields
    if (data.durationMinutes) {
      sanitized.durationMinutes = Math.max(0, Math.min(10080, Number(data.durationMinutes) || 0));
    }
    
    if (data.visitCountLocal) {
      sanitized.visitCountLocal = Math.max(1, Math.min(1000, Number(data.visitCountLocal) || 1));
    }
    
    return { ...data, ...sanitized };
  }
  
  static calculateDerivedFields(rowData) {
    const now = new Date();
    const firstVisit = rowData.firstVisitIso ? new Date(rowData.firstVisitIso) : null;
    const lastVisit = rowData.entryTime ? new Date(rowData.entryTime) : null;
    
    const accountAge = firstVisit ? Math.floor((now - firstVisit) / (1000 * 60 * 60 * 24)) : 0;
    const daysSinceLastVisit = lastVisit ? Math.floor((now - lastVisit) / (1000 * 60 * 60 * 24)) : 0;
    
    let userStatus = 'Inactive';
    if (accountAge <= 15) userStatus = 'New';
    else if (daysSinceLastVisit <= 30) userStatus = 'Active';
    
    let liveStatus = '🔴';
    if (rowData.entryTime && !rowData.exitTime) {
      const entryTime = new Date(rowData.entryTime);
      const hoursSinceEntry = (now - entryTime) / (1000 * 60 * 60);
      if (hoursSinceEntry <= 0.1) liveStatus = '🟢';
    } else if (daysSinceLastVisit <= 1) liveStatus = '🔵';
    else if (userStatus === 'New') liveStatus = '🆕';
    else if (userStatus === 'Active') liveStatus = '🟣';
    
    return {
      userStatus,
      liveStatus,
      accountAge,
      daysSinceLastVisit
    };
  }
}

/**
 * Web API: GET -> JSON export; POST -> tracking ingest
 */
function doGet(e) {
  const startTime = Date.now();
  Logger.info('GET request received', { parameters: e?.parameter });
  
  try {
    const format = (e?.parameter?.format || '').toLowerCase();
    const sheetId = e?.parameter?.sheetId || CONFIG.DEFAULT_SPREADSHEET_ID;
    const limit = Math.min(
      parseInt(e?.parameter?.limit || '', 10) || CONFIG.MAX_ROWS_EXPORT,
      CONFIG.MAX_ROWS_EXPORT
    );
    const uid = e?.parameter?.uid || null;

    const sheet = ensureSheet(sheetId);
    ensureFormulas(sheet);

    if (format === 'json') {
      const rows = readAllRowsAsObjects(sheet, { limit, uid });
      const responseTime = Date.now() - startTime;
      Logger.info('JSON export completed', { 
        rowCount: rows.length, 
        responseTime: `${responseTime}ms`,
        limit,
        uid: uid ? 'filtered' : 'all'
      });
      return jsonResponse({ 
        ok: true, 
        rows,
        meta: {
          totalRows: rows.length,
          responseTime: responseTime,
          timestamp: new Date().toISOString()
        }
      });
    }

    // Simple HTML info page if format not specified
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 50px auto; padding: 20px;">
        <h2>Users Dashboard Apps Script</h2>
        <p>✅ Service is running and operational</p>
        <p><strong>API Endpoints:</strong></p>
        <ul>
          <li>GET ?format=json - Export all data as JSON</li>
          <li>GET ?format=json&limit=100 - Export limited rows</li>
          <li>GET ?format=json&uid=USER_ID - Export specific user</li>
          <li>POST - Track page enter/exit events</li>
        </ul>
        <p><strong>Configuration:</strong></p>
        <ul>
          <li>Max Export Rows: ${CONFIG.MAX_ROWS_EXPORT}</li>
          <li>Cache Duration: ${CONFIG.CACHE_DURATION / 1000}s</li>
          <li>Log Level: ${CONFIG.LOG_LEVEL}</li>
        </ul>
      </div>
    `);
    return html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
  } catch (error) {
    Logger.error('GET request failed', { error: error.message, stack: error.stack });
    return jsonResponse({ 
      ok: false, 
      error: 'Internal server error',
      timestamp: new Date().toISOString()
    });
  }
}

function doPost(e) {
  const startTime = Date.now();
  Logger.info('POST request received');
  
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  
  try {
    lockAcquired = lock.tryLock(CONFIG.LOCK_TIMEOUT);
    if (!lockAcquired) {
      Logger.warn('Could not acquire lock, request may be queued');
      return jsonResponse({ 
        ok: false, 
        error: 'Service busy, please retry',
        retryAfter: 5
      });
    }
    
    const body = parsePostBody(e);
    if (!body) {
      Logger.warn('Invalid POST body received');
      return jsonResponse({ ok: false, error: 'No valid payload found' });
    }

    const sheetId = body.sheetId || CONFIG.DEFAULT_SPREADSHEET_ID;
    const sheet = ensureSheet(sheetId);
    ensureFormulas(sheet);

    const eventName = (body.event || '').toLowerCase();
    if (!body.uid) {
      Logger.warn('Missing UID in request', { event: eventName });
      return jsonResponse({ ok: false, error: 'Missing uid' });
    }

    // Validate and sanitize data
    const sanitizedBody = DataProcessor.sanitizeRowData(body);
    
    let res;
    switch (eventName) {
      case 'page_enter':
        const enterErrors = Validator.validatePageEnterData(sanitizedBody);
        if (enterErrors.length > 0) {
          Logger.warn('Page enter validation failed', { errors: enterErrors, uid: sanitizedBody.uid });
          return jsonResponse({ ok: false, error: 'Validation failed: ' + enterErrors.join(', ') });
        }
        res = handlePageEnter(sheet, sanitizedBody);
        break;
        
      case 'page_exit':
        const exitErrors = Validator.validatePageExitData(sanitizedBody);
        if (exitErrors.length > 0) {
          Logger.warn('Page exit validation failed', { errors: exitErrors, uid: sanitizedBody.uid });
          return jsonResponse({ ok: false, error: 'Validation failed: ' + exitErrors.join(', ') });
        }
        res = handlePageExit(sheet, sanitizedBody);
        break;
        
      default:
        Logger.warn('Unknown event type', { event: eventName, uid: sanitizedBody.uid });
        res = { ok: false, error: 'Unknown event: ' + eventName };
    }
    
    const responseTime = Date.now() - startTime;
    Logger.info('POST request completed', { 
      event: eventName, 
      uid: sanitizedBody.uid, 
      responseTime: `${responseTime}ms`,
      success: res?.ok
    });
    
    return jsonResponse(res || { ok: true });
    
  } catch (error) {
    Logger.error('POST request failed', { 
      error: error.message, 
      stack: error.stack,
      uid: body?.uid 
    });
    return jsonResponse({ 
      ok: false, 
      error: 'Internal server error',
      timestamp: new Date().toISOString()
    });
  } finally {
    if (lockAcquired) {
      try { 
        lock.releaseLock(); 
      } catch (releaseError) {
        Logger.warn('Failed to release lock', { error: releaseError.message });
      }
    }
  }
}

/**
 * Enhanced POST body parsing with better error handling
 */
function parsePostBody(e) {
  if (!e) return null;
  
  try {
    // x-www-form-urlencoded: payload=<json>
    if (e.parameter && e.parameter.payload) {
      const parsed = JSON.parse(e.parameter.payload);
      Logger.debug('Parsed form-encoded payload', { hasUid: !!parsed.uid, event: parsed.event });
      return parsed;
    }
    
    // raw JSON
    if (e.postData && e.postData.contents) {
      const parsed = JSON.parse(e.postData.contents);
      Logger.debug('Parsed raw JSON payload', { hasUid: !!parsed.uid, event: parsed.event });
      return parsed;
    }
    
    Logger.warn('No valid payload found in request');
    return null;
    
  } catch (parseError) {
    Logger.error('Failed to parse POST body', { 
      error: parseError.message,
      hasParameter: !!e.parameter,
      hasPostData: !!e.postData
    });
    return null;
  }
}

/**
 * Enhanced sheet management with caching
 */
function ensureSheet(spreadsheetId) {
  try {
    // Check cache first
    const now = Date.now();
    if (cache.spreadsheet && 
        cache.spreadsheet.getId() === spreadsheetId && 
        (now - cache.lastFetch) < CONFIG.CACHE_DURATION) {
      Logger.debug('Using cached spreadsheet');
      return cache.spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    }
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      Logger.info('Creating new sheet', { sheetName: CONFIG.SHEET_NAME });
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    }

    // Ensure header row
    const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    let needsHeader = false;
    for (let i = 0; i < HEADERS.length; i++) {
      if ((firstRow[i] || '') !== HEADERS[i]) { 
        needsHeader = true; 
        break; 
      }
    }
    
    if (needsHeader) {
      Logger.info('Updating sheet headers');
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
    
    // Update cache
    cache.spreadsheet = ss;
    cache.lastFetch = now;
    
    return sheet;
    
  } catch (error) {
    Logger.error('Failed to ensure sheet', { 
      spreadsheetId, 
      error: error.message 
    });
    throw error;
  }
}

/**
 * Enhanced formula management with better error handling
 */
function ensureFormulas(sheet) {
  try {
    Logger.debug('Ensuring formulas are set');
    
    const formulas = {
      'Days Since Last Visit': '=ARRAYFORMULA(IF(ROW(D2:D)=2, IF(D2<>"","",), IF(D2:D="", "", TODAY() - DATEVALUE(D2:D))))',
      'Account Age (Days)': '=ARRAYFORMULA(IF(ROW(C2:C)=2, IF(C2<>"","",), IF(C2:C="", "", TODAY() - DATEVALUE(C2:C))))',
      'User Status (Text)': [
        '=ARRAYFORMULA(IF(LEN(B2:B)=0, "",',
        ' IF(X2:X<=15, "New",',
        '  IF(W2:W>30, "Inactive", "Active")',
        ' )',
        '))'
      ].join(''),
      'Live Status Icon': [
        '=ARRAYFORMULA(IF(LEN(B2:B)=0, "",',
        ' IF( (LEN(S2:S)>0) * (LEN(T2:T)=0) * ((NOW()-S2:S) <= 0.1/24), "🟢",',
        '  IF(W2:W<=1, "🔵",',
        '    IF(V2:V="Inactive", "🔴",',
        '      IF(V2:V="New", "🆕", "🟣")',
        '    )',
        '  )',
        ' )',
        '))'
      ].join(''),
      'Repeat Visit Type': '=ARRAYFORMULA(IF(LEN(B2:B)=0, "", IF(H2:H<=1, "First", "Returning")))',
      'Total Sessions': '=ARRAYFORMULA(IF(LEN(B2:B)=0, "", H2:H))',
      'Active Period Range': [
        '=ARRAYFORMULA(IF(LEN(B2:B)=0, "",',
        ' IF(W2:W<=1,"≤24h",',
        '  IF(W2:W<=7,"≤7d",',
        '   IF(W2:W<=30,"≤30d",">30d")',
        '  )',
        ' )',
        '))'
      ].join('')
    };

    Object.entries(formulas).forEach(([header, formula]) => {
      setIfDifferent(sheet, 2, COL[header], formula);
    });
    
    Logger.debug('Formulas ensured successfully');
    
  } catch (error) {
    Logger.error('Failed to ensure formulas', { error: error.message });
    // Don't throw - formulas are not critical for basic functionality
  }
}

function setIfDifferent(sheet, row, col, formula) {
  try {
    const cell = sheet.getRange(row, col);
    const currentFormula = cell.getFormula();
    if (currentFormula !== formula) {
      cell.setFormula(formula);
      Logger.debug('Updated formula', { header: HEADERS[col-1] });
    }
  } catch (error) {
    Logger.warn('Failed to set formula', { 
      header: HEADERS[col-1], 
      error: error.message 
    });
  }
}

/**
 * Enhanced page enter handler
 */
function handlePageEnter(sheet, body) {
  const uid = String(body.uid || '').trim();
  const now = new Date();
  const entryIso = body.entryTime || now.toISOString();

  Logger.debug('Processing page enter', { uid, entryTime: entryIso });

  const rowIndex = findRowByUID(sheet, uid);
  if (rowIndex === -1) {
    // New user row
    Logger.info('Creating new user', { uid });
    
    const row = new Array(HEADERS.length).fill('');
    row[idx('Timestamp')] = now;
    row[idx('Unique User ID')] = uid;
    row[idx('First Visit DateTime')] = body.firstVisitIso ? new Date(body.firstVisitIso) : new Date(entryIso);
    row[idx('Last Visit DateTime')] = new Date(entryIso);
    row[idx('Last Exit DateTime')] = '';
    row[idx('Last Visited Duration (minutes)')] = 0;
    row[idx('Total Visited Duration (minutes)')] = 0;
    row[idx('Total Visit Count')] = 1;
    row[idx('Device Type')] = body.deviceType || '';
    row[idx('Device Model / Name')] = body.deviceModel || '';
    row[idx('Browser / WebView Info')] = body.userAgent || '';
    row[idx('Screen Resolution / Ratio')] = body.screenResolution || '';
    row[idx('Timezone')] = body.timezone || '';
    row[idx('IP Address')] = body.ip || '';
    row[idx('Country')] = body.country || '';
    row[idx('City')] = body.city || '';
    row[idx('Location (Latitude, Longitude)')] = body.location || '';
    row[idx('Last Visited Page / Section Name')] = body.sectionName || '';
    row[idx('App Entry Time')] = new Date(entryIso);
    row[idx('App Exit Time')] = '';
    row[idx('Source (optional)')] = body.source || '';
    row[idx('Notes / Admin Tag')] = '';

    sheet.appendRow(row);
    Logger.info('New user created successfully', { uid });
    return { ok: true, created: true, uid };
    
  } else {
    // Update existing user
    Logger.debug('Updating existing user', { uid, rowIndex });
    
    const rng = sheet.getRange(rowIndex, 1, 1, HEADERS.length);
    const vals = rng.getValues()[0];

    const prevTotalDur = toNumber(vals[idx('Total Visited Duration (minutes)')]);
    const prevVisitCount = toNumber(vals[idx('Total Visit Count')]);

    vals[idx('Timestamp')] = now;
    vals[idx('Last Visit DateTime')] = new Date(entryIso);
    vals[idx('App Entry Time')] = new Date(entryIso);
    vals[idx('App Exit Time')] = '';
    vals[idx('Last Visited Page / Section Name')] = body.sectionName || vals[idx('Last Visited Page / Section Name')] || '';

    // Update metadata (latest known)
    vals[idx('Device Type')] = body.deviceType || vals[idx('Device Type')] || '';
    vals[idx('Device Model / Name')] = body.deviceModel || vals[idx('Device Model / Name')] || '';
    vals[idx('Browser / WebView Info')] = body.userAgent || vals[idx('Browser / WebView Info')] || '';
    vals[idx('Screen Resolution / Ratio')] = body.screenResolution || vals[idx('Screen Resolution / Ratio')] || '';
    vals[idx('Timezone')] = body.timezone || vals[idx('Timezone')] || '';
    vals[idx('IP Address')] = body.ip || vals[idx('IP Address')] || '';
    vals[idx('Country')] = body.country || vals[idx('Country')] || '';
    vals[idx('City')] = body.city || vals[idx('City')] || '';
    vals[idx('Location (Latitude, Longitude)')] = body.location || vals[idx('Location (Latitude, Longitude)')] || '';
    vals[idx('Source (optional)')] = body.source || vals[idx('Source (optional)')] || '';

    // Counters
    vals[idx('Total Visit Count')] = prevVisitCount + 1;
    vals[idx('Last Visited Duration (minutes)')] = toNumber(vals[idx('Last Visited Duration (minutes)')]) || 0;
    vals[idx('Total Visited Duration (minutes)')] = prevTotalDur || 0;

    rng.setValues([vals]);
    Logger.info('User updated successfully', { uid, visitCount: prevVisitCount + 1 });
    return { ok: true, updated: true, uid };
  }
}

/**
 * Enhanced page exit handler
 */
function handlePageExit(sheet, body) {
  const uid = String(body.uid || '').trim();
  const now = new Date();
  const exitIso = body.exitTime || now.toISOString();
  const durMin = isFinite(body.durationMinutes) ? Number(body.durationMinutes) : 0;

  Logger.debug('Processing page exit', { uid, duration: durMin });

  let rowIndex = findRowByUID(sheet, uid);
  if (rowIndex === -1) {
    // If no existing row, create minimal and set exit
    Logger.info('Creating user on exit', { uid });
    
    const row = new Array(HEADERS.length).fill('');
    row[idx('Timestamp')] = now;
    row[idx('Unique User ID')] = uid;
    row[idx('First Visit DateTime')] = '';
    row[idx('Last Visit DateTime')] = '';
    row[idx('Last Exit DateTime')] = new Date(exitIso);
    row[idx('Last Visited Duration (minutes)')] = durMin;
    row[idx('Total Visited Duration (minutes)')] = durMin;
    row[idx('Total Visit Count')] = 1;
    row[idx('App Entry Time')] = '';
    row[idx('App Exit Time')] = new Date(exitIso);
    sheet.appendRow(row);
    
    Logger.info('User created on exit', { uid });
    return { ok: true, created: true, uid, note: 'Created on exit' };
    
  } else {
    Logger.debug('Updating user exit', { uid, rowIndex });
    
    const rng = sheet.getRange(rowIndex, 1, 1, HEADERS.length);
    const vals = rng.getValues()[0];

    const prevTotalDur = toNumber(vals[idx('Total Visited Duration (minutes)')]) || 0;

    vals[idx('Timestamp')] = now;
    vals[idx('Last Exit DateTime')] = new Date(exitIso);
    vals[idx('App Exit Time')] = new Date(exitIso);
    vals[idx('Last Visited Duration (minutes)')] = durMin;
    vals[idx('Total Visited Duration (minutes)')] = round2(prevTotalDur + durMin);

    rng.setValues([vals]);
    Logger.info('User exit updated', { uid, totalDuration: round2(prevTotalDur + durMin) });
    return { ok: true, updated: true, uid };
  }
}

/**
 * Enhanced data export with better performance
 */
function readAllRowsAsObjects(sheet, opts = {}) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.debug('No data rows found');
      return [];
    }
    
    const lastCol = HEADERS.length;
    const startRow = 2;
    const numRows = lastRow - 1;
    
    Logger.debug('Reading data', { 
      totalRows: numRows, 
      limit: opts.limit,
      uid: opts.uid 
    });

    const data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
    const rows = [];
    
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const uid = String(r[idx('Unique User ID')] || '').trim();
      if (!uid) continue;

      if (opts.uid && uid !== opts.uid) continue;

      const obj = {};
      for (let c = 0; c < HEADERS.length; c++) {
        const header = HEADERS[c];
        let val = r[c];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
        }
        obj[header] = val;
      }
      rows.push(obj);
      
      if (opts.limit && rows.length >= opts.limit) {
        Logger.debug('Reached limit', { limit: opts.limit });
        break;
      }
    }
    
    Logger.info('Data export completed', { 
      exportedRows: rows.length,
      totalAvailable: numRows,
      filtered: !!opts.uid
    });
    
    return rows;
    
  } catch (error) {
    Logger.error('Failed to read data', { 
      error: error.message,
      lastRow: sheet.getLastRow()
    });
    return [];
  }
}

/**
 * Enhanced helper functions
 */
function idx(header) { 
  return COL[header] - 1; // 0-based for row arrays
}

function toNumber(v) {
  if (v === '' || v == null) return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function round2(n) { 
  return Math.round(n * 100) / 100; 
}

function findRowByUID(sheet, uid) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return -1;
    
    const uidCol = COL['Unique User ID'];
    const range = sheet.getRange(2, uidCol, lastRow - 1, 1);
    const vals = range.getValues();
    
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '').trim() === uid) {
        return i + 2; // row index (1-based)
      }
    }
    return -1;
    
  } catch (error) {
    Logger.error('Failed to find row by UID', { 
      uid, 
      error: error.message 
    });
    return -1;
  }
}

function jsonResponse(obj) {
  try {
    const out = ContentService.createTextOutput(JSON.stringify(obj, null, 2));
    out.setMimeType(ContentService.MimeType.JSON);
    return out;
  } catch (error) {
    Logger.error('Failed to create JSON response', { error: error.message });
    const errorResponse = { ok: false, error: 'Response generation failed' };
    const out = ContentService.createTextOutput(JSON.stringify(errorResponse));
    out.setMimeType(ContentService.MimeType.JSON);
    return out;
  }
}

/**
 * Utility function to clear cache (useful for testing)
 */
function clearCache() {
  cache.sheetData = null;
  cache.lastFetch = 0;
  cache.spreadsheet = null;
  Logger.info('Cache cleared');
}

/**
 * Health check function
 */
function healthCheck() {
  try {
    const sheet = ensureSheet(CONFIG.DEFAULT_SPREADSHEET_ID);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    return {
      ok: true,
      timestamp: new Date().toISOString(),
      sheetInfo: {
        name: CONFIG.SHEET_NAME,
        lastRow,
        lastCol,
        headers: HEADERS.length
      },
      config: {
        maxRowsExport: CONFIG.MAX_ROWS_EXPORT,
        cacheDuration: CONFIG.CACHE_DURATION,
        logLevel: CONFIG.LOG_LEVEL
      }
    };
  } catch (error) {
    Logger.error('Health check failed', { error: error.message });
    return {
      ok: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}