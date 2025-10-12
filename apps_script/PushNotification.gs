// PushNotification.gs - Stubs for push workflow (to be wired with FCM/Web Push)

function sendPush(deviceToken, title, message) {
  // Placeholder for integration with FCM or Web Push service
  // Log intent in a sheet named PushLogs
  const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  let sheet = ss.getSheetByName('PushLogs');
  if (!sheet) sheet = ss.insertSheet('PushLogs');
  if (sheet.getLastRow() === 0) sheet.appendRow(['timestamp_utc','token','title','message','status']);
  sheet.appendRow([new Date().toISOString(), deviceToken, title, message, 'queued']);
}

function broadcastPush(title, message) {
  // Read distinct device tokens from Master sheet
  const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const master = ss.getSheetByName('Master');
  if (!master) return 0;
  const values = master.getDataRange().getValues();
  const headers = values.shift();
  const tokenIdx = headers.indexOf('device_token'); // if exists
  if (tokenIdx === -1) return 0;
  const tokens = Array.from(new Set(values.map(r => r[tokenIdx]).filter(Boolean)));
  tokens.forEach(t => sendPush(t, title, message));
  return tokens.length;
}

function scheduleDailyInactivePush() {
  // Example time-driven trigger setup (run once to set trigger)
  ScriptApp.newTrigger('runDailyInactivePush')
    .timeBased().atHour(23).everyDays(1).create();
}

function runDailyInactivePush() {
  // Scan for users inactive >24h and send default push
  const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const master = ss.getSheetByName('Master');
  if (!master) return;
  const values = master.getDataRange().getValues();
  const headers = values.shift();
  const userIdx = headers.indexOf('user_id');
  const nameIdx = headers.indexOf('user_name');
  const lastIdx = headers.indexOf('status_last_updated');
  const tokenIdx = headers.indexOf('device_token');
  const now = new Date();
  const byUser = {};
  values.forEach(r => {
    const uid = r[userIdx];
    const last = new Date(r[lastIdx] || 0);
    const token = tokenIdx >= 0 ? r[tokenIdx] : '';
    if (!byUser[uid] || new Date(byUser[uid].last) < last) {
      byUser[uid] = { last, token, name: r[nameIdx] };
    }
  });
  const dayMs = 24*60*60*1000;
  const title = 'We missed you today!';
  const message = 'Come back and check what\'s new.';
  Object.values(byUser).forEach(u => {
    if (u.last && (now - u.last) > dayMs && u.token) {
      sendPush(u.token, title, message);
    }
  });
}
