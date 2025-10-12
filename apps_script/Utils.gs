// Utils.gs - Helpers for Apps Script backend

const SECRET_TOKEN = 'YOUR_SECRET_KEY';

const Utils = {
  validateToken(token) {
    return SECRET_TOKEN ? token === SECRET_TOKEN : true;
  },

  jsonResponse(obj, code = 200) {
    return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON)
      .setResponseCode(code);
  },

  getOrCreateTodaySheet(spreadsheetId, headers) {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const name = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
    }
    this.ensureHeaders(sheet, headers);
    return sheet;
  },

  ensureHeaders(sheet, headers) {
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const needs = headers.some((h, i) => firstRow[i] !== h);
    if (needs) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  },

  appendToMasterIfExists(spreadsheetId, rows, headers) {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let master = ss.getSheetByName('Master');
    if (!master) return;
    if (master.getLastRow() === 0) master.appendRow(headers);
    master.getRange(master.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  },

  getSettings() {
    const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Settings');
    if (!sheet) {
      sheet = ss.insertSheet('Settings');
      sheet.appendRow(['key', 'value']);
      sheet.appendRow(['auto_push_enabled', 'false']);
      sheet.appendRow(['auto_push_hour', '23']);
      sheet.appendRow(['auto_push_minute', '0']);
      sheet.appendRow(['default_message', 'We missed you today! Come back and check what\'s new.']);
    }
    const values = sheet.getDataRange().getValues();
    const obj = {};
    values.slice(1).forEach(r => obj[String(r[0])] = String(r[1]));
    return {
      auto_push_enabled: obj.auto_push_enabled === 'true',
      auto_push_hour: Number(obj.auto_push_hour || 23),
      auto_push_minute: Number(obj.auto_push_minute || 0),
      default_message: obj.default_message || ''
    };
  },

  setSettings(settings) {
    const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Settings');
    if (!sheet) {
      sheet = ss.insertSheet('Settings');
      sheet.appendRow(['key', 'value']);
    }
    const map = this.getSettings();
    const merged = Object.assign({}, map, settings);
    const data = [
      ['auto_push_enabled', String(merged.auto_push_enabled)],
      ['auto_push_hour', String(merged.auto_push_hour)],
      ['auto_push_minute', String(merged.auto_push_minute)],
      ['default_message', String(merged.default_message)]
    ];
    sheet.clearContents();
    sheet.appendRow(['key', 'value']);
    sheet.getRange(2, 1, data.length, 2).setValues(data);
  },

  payloadToRow(p, now = new Date()) {
    const tsUtc = Utilities.formatDate(now, 'Etc/UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
    const localDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const localTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

    const row = [
      tsUtc,
      localDate,
      localTime,
      p.event_type || '',
      p.event_label || '',
      p.session_id || '',
      p.user_id || '',
      p.user_name || 'Guest User',
      p.is_repeated_user === true ? true : false,
      Number(p.visit_count_total || 0),
      Number(p.visit_count_today || 0),
      p.session_start_time || '',
      p.session_end_time || '',
      Number(p.session_active_seconds || 0),
      Number(p.session_duration_minutes || 0),
      Number(p.session_reopen_count || 0),
      Number(p.idle_time_seconds || 0),
      p.status_online || '',
      p.status_last_updated || '',
      Number(p.total_session_time_user || 0),
      p.android_id || '',
      p.advertising_id || '',
      p.app_generated_uuid || '',
      p.device_name || '',
      p.device_model || '',
      p.device_fingerprint_hash || '',
      p.unique_user_signature || '',
      p.device_os || '',
      p.os_version || '',
      p.browser_name || '',
      p.browser_version || '',
      p.user_agent || '',
      Number(p.screen_width || 0),
      Number(p.screen_height || 0),
      p.screen_ratio || '',
      p.orientation || '',
      Number(p.pixel_density || 0),
      Number(p.color_depth || 0),
      p.public_ip || '',
      p.ip_city || '',
      p.ip_region || '',
      p.ip_country || '',
      Number(p.ip_latitude || 0),
      Number(p.ip_longitude || 0),
      p.geo_location_text || '',
      p.timezone || '',
      p.network_type || '',
      p.connection_speed_est || '',
      p.referrer || '',
      p.page_url || '',
      p.redirect_target || '',
      p.app_version || '',
      p.platform_type || '',
      Number(p.battery_level ?? ''),
      p.is_charging === true ? true : false,
      p.language || '',
      p.theme_mode || '',
      p.screen_focus_state || '',
      Number(p.heartbeat_count || 0),
      p.offline_buffered === true ? true : false,
      p.sync_time || '',
      p.token || '',
      p.error_log || '',
      p.note || '',
      JSON.stringify(p.extra_json || {})
    ];
    return row;
  }
};
