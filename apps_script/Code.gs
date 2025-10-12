// Code.gs - Entry point for web app and event logging

const MASTER_SPREADSHEET_ID = '11trONTNEuQ4WfmzTynytufAsBhYzgCkCem3L5w2VUvo';
const SHEET_HEADERS = [
  'timestamp_utc','date_local','time_local','event_type','event_label','session_id','user_id','user_name','is_repeated_user','visit_count_total','visit_count_today','session_start_time','session_end_time','session_active_seconds','session_duration_minutes','session_reopen_count','idle_time_seconds','status_online','status_last_updated','total_session_time_user','android_id','advertising_id','app_generated_uuid','device_name','device_model','device_fingerprint_hash','unique_user_signature','device_os','os_version','browser_name','browser_version','user_agent','screen_width','screen_height','screen_ratio','orientation','pixel_density','color_depth','public_ip','ip_city','ip_region','ip_country','ip_latitude','ip_longitude','geo_location_text','timezone','network_type','connection_speed_est','referrer','page_url','redirect_target','app_version','platform_type','battery_level','is_charging','language','theme_mode','screen_focus_state','heartbeat_count','offline_buffered','sync_time','token','error_log','note','extra_json'
];

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = (params.action || '').toLowerCase();
  try {
    if (action === 'dashboard') {
      return ContentService.createTextOutput(JSON.stringify({ ok: true, data: getTodayData() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'settings') {
      const settings = Utils.getSettings();
      return ContentService.createTextOutput(JSON.stringify({ ok: true, settings }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput('OK');
  } catch (err) {
    return Utils.jsonResponse({ ok: false, error: String(err) }, 500);
  }
}

function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents || '{}');
    // Admin actions routing
    if (json.admin_action) {
      if (!Utils.validateToken(json.token)) {
        return Utils.jsonResponse({ ok: false, error: 'invalid_token' }, 403);
      }
      const a = String(json.admin_action).toLowerCase();
      if (a === 'broadcast') {
        const count = broadcastPush(json.title || '', json.message || '');
        return Utils.jsonResponse({ ok: true, sent: count });
      }
      if (a === 'send') {
        sendPush(json.device_token || '', json.title || '', json.message || '');
        return Utils.jsonResponse({ ok: true });
      }
      if (a === 'set_settings') {
        Utils.setSettings({
          auto_push_enabled: Boolean(json.auto_push_enabled),
          auto_push_hour: Number(json.auto_push_hour || 23),
          auto_push_minute: Number(json.auto_push_minute || 0),
          default_message: String(json.default_message || 'We missed you today! Come back and check what\'s new.')
        });
        return Utils.jsonResponse({ ok: true });
      }
      return Utils.jsonResponse({ ok: false, error: 'unknown_admin_action' }, 400);
    }

    if (!Utils.validateToken(json.token)) {
      return Utils.jsonResponse({ ok: false, error: 'invalid_token' }, 403);
    }
    const payloads = Array.isArray(json) ? json : (Array.isArray(json.payloads) ? json.payloads : [json]);
    const now = new Date();
    const sheet = Utils.getOrCreateTodaySheet(MASTER_SPREADSHEET_ID, SHEET_HEADERS);
    const rows = payloads.map(p => Utils.payloadToRow(p, now));
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, SHEET_HEADERS.length).setValues(rows);
    // Mirror to master summary tab if exists
    Utils.appendToMasterIfExists(MASTER_SPREADSHEET_ID, rows, SHEET_HEADERS);
    return Utils.jsonResponse({ ok: true, inserted: rows.length });
  } catch (err) {
    return Utils.jsonResponse({ ok: false, error: String(err) }, 500);
  }
}
