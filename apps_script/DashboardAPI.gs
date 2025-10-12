// DashboardAPI.gs - Simple endpoints to read data for dashboards

function getTodayData() {
  const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const name = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map(r => headers.reduce((o, h, i) => { o[h] = r[i]; return o; }, {}));
}

function doGetDashboard(e) {
  const data = getTodayData();
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}
