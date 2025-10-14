/**
 * Google Apps Script backend for Users Analytics
 * Sheet Name must be exactly: "Users Dashboard"
 * Spreadsheet ID provided by user.
 */

const SPREADSHEET_ID = '11trONTNEuQ4WfmzTynytufAsBhYzgCkCem3L5w2VUvo';
const SHEET_NAME = 'Users Dashboard';

// Column indices (1-based)
const COLS = {
  Timestamp: 1,
  UID: 2,
  FirstVisit: 3,
  LastVisit: 4,
  LastExit: 5,
  LastDurationMin: 6,
  TotalDurationMin: 7,
  TotalVisitCount: 8,
  DeviceType: 9,
  DeviceModel: 10,
  UserAgent: 11,
  Screen: 12,
  Timezone: 13,
  IP: 14,
  Country: 15,
  City: 16,
  Location: 17,
  LastPage: 18,
  EntryTime: 19,
  ExitTime: 20,
  LiveIcon: 21,
  UserStatus: 22,
  DaysSinceLast: 23,
  AccountAgeDays: 24,
  RepeatVisitType: 25,
  TotalSessions: 26,
  ActivePeriodRange: 27,
  Source: 28,
  Notes: 29,
};

/** Ensure sheet and headers exist */
function ensureSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  const headers = [
    'Timestamp','Unique User ID','First Visit DateTime','Last Visit DateTime','Last Exit DateTime',
    'Last Visited Duration (minutes)','Total Visited Duration (minutes)','Total Visit Count',
    'Device Type','Device Model / Name','Browser / WebView Info','Screen Resolution / Ratio','Timezone',
    'IP Address','Country','City','Location (Latitude, Longitude)','Last Visited Page / Section Name',
    'App Entry Time','App Exit Time','Live Status Icon','User Status (Text)','Days Since Last Visit','Account Age (Days)',
    'Repeat Visit Type','Total Sessions','Active Period Range','Source (optional)','Notes / Admin Tag'
  ];

  const r1 = sh.getRange(1,1,1,headers.length);
  const hasHeader = r1.getValues()[0].filter(String).length > 0;
  if (!hasHeader) {
    r1.setValues([headers]);
    sh.getFrozenRows() === 0 && sh.setFrozenRows(1);
  }
  return sh;
}

/** Main entry for POST events from client */
function doPost(e) {
  try {
    const sh = ensureSheet_();
    const payload = parseRequestBody_(e);
    if (!payload || !payload.uid) return json_({ ok:false, error:'Missing uid' });

    const lock = LockService.getScriptLock();
    lock.tryLock(5000);
    try {
      const rowIndex = findOrCreateRowByUid_(sh, payload.uid, payload);

      if (payload.eventType === 'enter') handleEnter_(sh, rowIndex, payload);
      else if (payload.eventType === 'exit') handleExit_(sh, rowIndex, payload);
      else handleUpsert_(sh, rowIndex, payload);

      return json_({ ok:true });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

/** Returns dashboard HTML or JSON data depending on params */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  if (params && (params.format === 'json' || params.format === 'JSON')) {
    return json_(getUsersData_());
  }
  // Default: serve dashboard UI (single-file HTML)
  const html = HtmlService.createHtmlOutput(DASHBOARD_HTML_())
    .setTitle('Users Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

/** Provide JSON to google.script.run from the dashboard */
function getUsersData() {
  return getUsersData_();
}

/** Internal: read all rows into objects */
function getUsersData_() {
  const sh = ensureSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows: [] };
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  const rows = values.map(r => ({
    Timestamp: r[COLS.Timestamp-1],
    UID: r[COLS.UID-1],
    FirstVisit: r[COLS.FirstVisit-1],
    LastVisit: r[COLS.LastVisit-1],
    LastExit: r[COLS.LastExit-1],
    LastDurationMin: r[COLS.LastDurationMin-1],
    TotalDurationMin: r[COLS.TotalDurationMin-1],
    TotalVisitCount: r[COLS.TotalVisitCount-1],
    DeviceType: r[COLS.DeviceType-1],
    DeviceModel: r[COLS.DeviceModel-1],
    UserAgent: r[COLS.UserAgent-1],
    Screen: r[COLS.Screen-1],
    Timezone: r[COLS.Timezone-1],
    IP: r[COLS.IP-1],
    Country: r[COLS.Country-1],
    City: r[COLS.City-1],
    Location: r[COLS.Location-1],
    LastPage: r[COLS.LastPage-1],
    EntryTime: r[COLS.EntryTime-1],
    ExitTime: r[COLS.ExitTime-1],
    LiveIcon: r[COLS.LiveIcon-1],
    UserStatus: r[COLS.UserStatus-1],
    DaysSinceLast: r[COLS.DaysSinceLast-1],
    AccountAgeDays: r[COLS.AccountAgeDays-1],
    RepeatVisitType: r[COLS.RepeatVisitType-1],
    TotalSessions: r[COLS.TotalSessions-1],
    ActivePeriodRange: r[COLS.ActivePeriodRange-1],
    Source: r[COLS.Source-1],
    Notes: r[COLS.Notes-1],
  }));
  return { rows };
}

/** Helpers */
function parseRequestBody_(e) {
  if (!e) return null;
  if (e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch(_) {}
  }
  // Also accept urlencoded fallback
  const params = e.parameter || {};
  if (params.json) { try { return JSON.parse(params.json); } catch(_) {} }
  return params || null;
}

function findOrCreateRowByUid_(sh, uid, payload) {
  const data = sh.getRange(2, COLS.UID, Math.max(0, sh.getLastRow()-1), 1).getValues();
  for (let i=0; i<data.length; i++) {
    if (String(data[i][0]).trim() === String(uid).trim()) {
      return i + 2; // row index
    }
  }
  // Create new row
  const rowIndex = sh.getLastRow() + 1;
  const now = new Date();
  const firstVisit = payload && payload.firstVisitDateTime ? new Date(payload.firstVisitDateTime) : now;
  sh.getRange(rowIndex, 1, 1, 29).setValues([[
    formatDate_(now),                 // Timestamp (A)
    uid,                              // B UID
    formatDate_(firstVisit),          // C First Visit
    formatDate_(now),                 // D Last Visit (init)
    '',                               // E Last Exit
    '',                               // F Last Duration
    '',                               // G Total Duration
    0,                                // H Total Visit Count (we will bump in enter)
    '', '', '', '', '',               // I..M Device/UA/Screen/TZ (filled later)
    '', '', '',                       // N..P IP/Country/City
    '',                               // Q Location
    '',                               // R Last Page
    '',                               // S Entry
    '',                               // T Exit
    '🆕',                             // U Live Icon (New)
    'New',                            // V Status
    0,                                // W Days Since Last
    0,                                // X Account Age Days
    '',                               // Y Repeat Visit Type
    0,                                // Z Total Sessions
    '',                               // AA Active Period Range
    payload && payload.source ? payload.source : '', // AB Source
    ''                                // AC Notes
  ]]);
  return rowIndex;
}

function handleEnter_(sh, rowIndex, p) {
  const now = new Date();
  const row = sh.getRange(rowIndex, 1, 1, 29).getValues()[0];

  const firstVisitStr = row[COLS.FirstVisit-1];
  const firstVisit = firstVisitStr ? new Date(firstVisitStr) : now;

  const prevCount = Number(row[COLS.TotalVisitCount-1]) || 0;
  const newCount = Math.max(prevCount, 0) + 1;

  // Update device and env info each enter
  const updates = [];
  const indices = [];
  function set(col, val){ updates.push(val); indices.push(col); }

  set(COLS.Timestamp, formatDate_(now));
  set(COLS.LastVisit, formatDate_(now));
  set(COLS.TotalVisitCount, newCount);
  set(COLS.TotalSessions, newCount);

  if (p.deviceType) set(COLS.DeviceType, p.deviceType);
  if (p.deviceModel) set(COLS.DeviceModel, p.deviceModel);
  if (p.userAgent) set(COLS.UserAgent, p.userAgent);
  if (p.screen) set(COLS.Screen, p.screen);
  if (p.timezone) set(COLS.Timezone, p.timezone);
  if (p.ip) set(COLS.IP, p.ip);
  if (p.country) set(COLS.Country, p.country);
  if (p.city) set(COLS.City, p.city);
  if (p.location) set(COLS.Location, p.location);
  if (p.lastVisitedPage) set(COLS.LastPage, p.lastVisitedPage);
  if (p.appEntryTime) set(COLS.EntryTime, formatDate_(new Date(p.appEntryTime)));
  if (p.source) set(COLS.Source, p.source);

  // Status logic
  const accAgeDays = daysBetween_(firstVisit, now);
  const daysSinceLast = 0; // just visited
  const icon = accAgeDays <= 15 ? '🆕' : '🟢'; // live when entering; new for first 15 days
  const statusText = accAgeDays <= 15 ? 'New' : 'Live';

  set(COLS.LiveIcon, icon);
  set(COLS.UserStatus, statusText);
  set(COLS.DaysSinceLast, daysSinceLast);
  set(COLS.AccountAgeDays, accAgeDays);
  set(COLS.ActivePeriodRange, 'Last 24hr');

  // Apply updates in one batch
  if (indices.length) writeSparse_(sh, rowIndex, indices, updates);
}

function handleExit_(sh, rowIndex, p) {
  const now = new Date();
  const row = sh.getRange(rowIndex, 1, 1, 29).getValues()[0];

  const firstVisitStr = row[COLS.FirstVisit-1];
  const firstVisit = firstVisitStr ? new Date(firstVisitStr) : now;
  const lastVisitStr = row[COLS.LastVisit-1];
  const lastVisit = lastVisitStr ? new Date(lastVisitStr) : now;

  const prevTotalDur = Number(row[COLS.TotalDurationMin-1]) || 0;
  const lastDur = Number(p.lastVisitedDurationMinutes) || 0;
  const newTotalDur = Math.round((prevTotalDur + Math.max(0, lastDur)) * 100) / 100;

  const updates = [];
  const indices = [];
  function set(col, val){ updates.push(val); indices.push(col); }

  set(COLS.Timestamp, formatDate_(now));
  set(COLS.LastExit, p.lastExitDateTime ? formatDate_(new Date(p.lastExitDateTime)) : formatDate_(now));
  if (p.appExitTime) set(COLS.ExitTime, formatDate_(new Date(p.appExitTime)));
  if (p.appEntryTime) set(COLS.EntryTime, formatDate_(new Date(p.appEntryTime)));
  set(COLS.LastDurationMin, lastDur);
  set(COLS.TotalDurationMin, newTotalDur);

  if (p.deviceType) set(COLS.DeviceType, p.deviceType);
  if (p.deviceModel) set(COLS.DeviceModel, p.deviceModel);
  if (p.userAgent) set(COLS.UserAgent, p.userAgent);
  if (p.screen) set(COLS.Screen, p.screen);
  if (p.timezone) set(COLS.Timezone, p.timezone);
  if (p.ip) set(COLS.IP, p.ip);
  if (p.country) set(COLS.Country, p.country);
  if (p.city) set(COLS.City, p.city);
  if (p.location) set(COLS.Location, p.location);
  if (p.lastVisitedPage) set(COLS.LastPage, p.lastVisitedPage);

  // Status after exit
  const daysSinceLast = daysBetween_(lastVisit, now) <= 1 ? 0 : daysBetween_(lastVisit, now);
  const accAgeDays = daysBetween_(firstVisit, now);

  let icon = '🟣';
  let statusText = 'Active';

  if (accAgeDays <= 15) { icon = '🆕'; statusText = 'New'; }
  else if (daysSinceLast <= 1) { icon = '🔵'; statusText = 'Active (24h)'; }
  else if (daysSinceLast <= 30) { icon = '🟣'; statusText = 'Active'; }
  else { icon = '🔴'; statusText = 'Inactive'; }

  set(COLS.LiveIcon, icon);
  set(COLS.UserStatus, statusText);
  set(COLS.DaysSinceLast, daysSinceLast);
  set(COLS.AccountAgeDays, accAgeDays);
  set(COLS.ActivePeriodRange, daysSinceLast <= 1 ? 'Last 24hr' : (daysSinceLast <= 7 ? 'Last 7 days' : (daysSinceLast <= 30 ? 'Last 30 days' : 'Older')));

  // Repeat Visit Type heuristic
  const totalVisits = Number(row[COLS.TotalVisitCount-1]) || 0;
  const ageDays = Math.max(accAgeDays, 1);
  const visitsPerDay = totalVisits / ageDays;
  let repeatType = 'Rare';
  if (visitsPerDay >= 0.8) repeatType = 'Daily';
  else if (visitsPerDay >= 0.2) repeatType = 'Weekly';
  else if (visitsPerDay >= 0.05) repeatType = 'Monthly';
  set(COLS.RepeatVisitType, repeatType);

  if (indices.length) writeSparse_(sh, rowIndex, indices, updates);
}

function handleUpsert_(sh, rowIndex, p) {
  // generic upsert with minimal fields
  const now = new Date();
  const updates = [];
  const indices = [];
  function set(col, val){ updates.push(val); indices.push(col); }
  set(COLS.Timestamp, formatDate_(now));
  if (p.lastVisitedPage) set(COLS.LastPage, p.lastVisitedPage);
  if (p.ip) set(COLS.IP, p.ip);
  if (p.city) set(COLS.City, p.city);
  if (p.country) set(COLS.Country, p.country);
  if (indices.length) writeSparse_(sh, rowIndex, indices, updates);
}

function writeSparse_(sh, rowIndex, indices, values) {
  // Batch update sparse columns in one call
  const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  for (let i=0; i<indices.length; i++) {
    const col = indices[i];
    row[col-1] = values[i];
  }
  sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function formatDate_(d) {
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm:ss'); } catch(_) { return ''; }
}

function daysBetween_(a, b) {
  const ms = Math.abs(new Date(b).getTime() - new Date(a).getTime());
  return Math.floor(ms / 86400000);
}

/**
 * Single-file Dashboard HTML returned by doGet (no external files).
 * If you prefer to use a standalone HTML file, you can paste the same markup from UsersDashboard.html.
 */
function DASHBOARD_HTML_(){
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Users Dashboard</title>
<style>
  :root{
    --bg: #070b18; --panel: #0d1328; --fg:#eaf0ff; --muted:#96a3c7; --primary:#6c8cff; --ok:#1ec97f; --warn:#ffcc66; --err:#ff5a78; --live:#37b5ff; --card:#0f1733; --border:rgba(255,255,255,0.08);
    --radius:14px; --shadow:0 10px 30px rgba(0,0,0,0.25);
    font-size:16px;
  }
  @media (max-width: 900px){ :root{ font-size:15px; } }
  @media (max-width: 600px){ :root{ font-size:14px; } }
  @media (max-width: 420px){ :root{ font-size:13.5px; } }

  *{ box-sizing:border-box; }
  body{ margin:0; background:linear-gradient(180deg,#070b18,#0a1021); color:var(--fg); font-family: system-ui, -apple-system, Segoe UI, Roboto, Ubuntu, Cantarell, Noto Sans, Arial; }

  header{ padding:1rem 1.25rem; position:sticky; top:0; background:rgba(7,11,24,0.8); backdrop-filter: blur(10px); border-bottom:1px solid var(--border); z-index:5; }
  .title{ display:flex; align-items:center; justify-content:space-between; gap:1rem; }
  .title h1{ margin:0; font-size:1.25rem; }
  .subtitle{ margin:0.25rem 0 0 0; color:var(--muted); font-size:0.9rem; }

  main{ padding:1rem; display:grid; grid-template-columns: 1fr; gap:1rem; max-width: 1400px; margin: 0 auto; }
  .grid{ display:grid; grid-template-columns: repeat(12, 1fr); gap:1rem; }
  .card{ background:linear-gradient(180deg,#0d1328,#0b1122); border:1px solid var(--border); border-radius:var(--radius); box-shadow:var(--shadow); padding:1rem; min-height:auto; }
  .kpis{ grid-column: 1/-1; display:grid; grid-template-columns: repeat(6,1fr); gap:1rem; }
  @media (max-width: 1100px){ .kpis{ grid-template-columns: repeat(3,1fr); } }
  @media (max-width: 600px){ .kpis{ grid-template-columns: repeat(2,1fr); } }

  .kpi{ position:relative; overflow:hidden; }
  .kpi h3{ margin:0; font-size:0.9rem; color:var(--muted); }
  .kpi .val{ font-size:2rem; font-weight:800; letter-spacing:0.5px; margin-top:0.15rem; }
  .kpi .spark{ position:absolute; right:0.75rem; bottom:0.5rem; width:100px; height:32px; opacity:0.8; }
  .kpi .bubble{ position:absolute; inset:auto -20px -30px auto; width:140px; height:140px; border-radius:50%; filter: blur(22px); opacity:0.22; }
  .kpi.total .bubble{ background: var(--primary); }
  .kpi.active .bubble{ background: var(--ok); }
  .kpi.inactive .bubble{ background: var(--err); }
  .kpi.today .bubble{ background: var(--live); }
  .kpi.live .bubble{ background: var(--live); }
  .kpi.last24 .bubble{ background: var(--warn); }

  .filters{ grid-column: 1/-1; display:flex; flex-wrap:wrap; gap:0.75rem; align-items:center; }
  .filters .f{ display:flex; align-items:center; gap:0.5rem; padding:0.5rem 0.75rem; background:var(--card); border:1px solid var(--border); border-radius:999px; }
  .filters select, .filters input{ background:transparent; color:var(--fg); border:none; outline:none; font-size:0.95rem; min-width: 8rem; }

  .lists{ grid-column: 1/-1; display:grid; grid-template-columns: 2fr 1fr; gap:1rem; }
  @media (max-width: 1000px){ .lists{ grid-template-columns: 1fr; } }

  .table{ width:100%; border-collapse: collapse; font-size:0.9rem; }
  .table th, .table td{ border-bottom:1px dashed var(--border); padding:0.6rem 0.5rem; text-align:left; vertical-align: top; }
  .table th{ color:var(--muted); font-weight:600; }
  .tag{ display:inline-flex; align-items:center; gap:0.4rem; padding:0.2rem 0.5rem; border-radius: 0.5rem; background:rgba(255,255,255,0.04); border: 1px solid var(--border); }
  .tag i{ font-style: normal; }

  .chart{ height: 220px; width:100%; }

  .muted{ color: var(--muted); }
</style>
</head>
<body>
<header>
  <div class="title">
    <div>
      <h1>Users Dashboard</h1>
      <p class="subtitle">Live, Active, Inactive, New — with filters and charts</p>
    </div>
    <div class="tag"><i>🗂️</i> Sheet: Users Dashboard</div>
  </div>
</header>
<main>
  <section class="grid">
    <div class="kpis">
      <div class="card kpi total"><h3>Total Users</h3><div class="val" id="k_total">0</div><canvas class="spark" id="sp_total"></canvas><div class="bubble"></div></div>
      <div class="card kpi active"><h3>Active (≤30d)</h3><div class="val" id="k_active">0</div><canvas class="spark" id="sp_active"></canvas><div class="bubble"></div></div>
      <div class="card kpi inactive"><h3>Inactive (>30d)</h3><div class="val" id="k_inactive">0</div><canvas class="spark" id="sp_inactive"></canvas><div class="bubble"></div></div>
      <div class="card kpi today"><h3>Today's Active</h3><div class="val" id="k_today">0</div><canvas class="spark" id="sp_today"></canvas><div class="bubble"></div></div>
      <div class="card kpi live"><h3>Current Live</h3><div class="val" id="k_live">0</div><canvas class="spark" id="sp_live"></canvas><div class="bubble"></div></div>
      <div class="card kpi last24"><h3>Last 24h Active</h3><div class="val" id="k_24h">0</div><canvas class="spark" id="sp_24h"></canvas><div class="bubble"></div></div>
    </div>

    <div class="card filters">
      <div class="f"><span class="muted">Status</span>
        <select id="f_status">
          <option value="all">All</option>
          <option value="live">Live</option>
          <option value="active">Active</option>
          <option value="inactive">Inactive</option>
          <option value="new">New</option>
        </select>
      </div>
      <div class="f"><span class="muted">Time Range</span>
        <select id="f_range">
          <option value="all">All time</option>
          <option value="24h">Last 24h</option>
          <option value="3d">Last 3 days</option>
          <option value="7d">Last 7 days</option>
          <option value="30d">Last 30 days</option>
          <option value="90d">Last 3 months</option>
          <option value="180d">Last 6 months</option>
          <option value="365d">Last 12 months</option>
        </select>
      </div>
      <div class="f"><span class="muted">Country</span>
        <input id="f_country" placeholder="Any" />
      </div>
      <div class="f"><span class="muted">City</span>
        <input id="f_city" placeholder="Any" />
      </div>
      <div class="f"><span class="muted">Min Minutes</span>
        <input id="f_minmin" type="number" min="0" step="1" placeholder="0" />
      </div>
      <div class="f"><span class="muted">Search</span>
        <input id="f_search" placeholder="UID / Page / IP" />
      </div>
    </div>

    <div class="lists">
      <div class="card">
        <h3>Users List</h3>
        <table class="table" id="tbl">
          <thead>
            <tr>
              <th>User Id</th><th>Timestamp</th><th>IP</th><th>City</th><th>Country</th>
              <th>Last Visit</th><th>Last Duration (m)</th><th>Total Duration (m)</th><th>Icon</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>
      <div class="card">
        <h3>All-time Totals</h3>
        <div class="muted" id="totals_text">Loading...</div>
        <canvas class="chart" id="chart_totals"></canvas>
      </div>
    </div>
  </section>
</main>
<script>
(function(){
  const byId = (id) => document.getElementById(id);
  const F = {
    status: byId('f_status'), range: byId('f_range'), country: byId('f_country'), city: byId('f_city'), minmin: byId('f_minmin'), search: byId('f_search')
  };

  let ROWS = [];

  function parseDate(s){ if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d; }
  function daysBetween(a,b){ if(!a||!b) return Infinity; return Math.floor(Math.abs(b - a)/86400000); }

  function normalize(r){
    const now = new Date();
    const first = parseDate(r.FirstVisit);
    const last = parseDate(r.LastVisit);

    const daysSince = last ? daysBetween(last, now) : Infinity;
    const age = first ? daysBetween(first, now) : 0;

    // Derive status
    let status = 'inactive', icon = '🔴';
    if (age <= 15) { status = 'new'; icon = '🆕'; }
    else if (daysSince <= 0) { status = 'live'; icon = '🟢'; }
    else if (daysSince <= 1) { status = 'active-24h'; icon = '🔵'; }
    else if (daysSince <= 30) { status = 'active'; icon = '🟣'; }

    return {
      uid: r.UID || r["Unique User ID"] || '',
      timestamp: r.Timestamp || '',
      ip: r.IP || '',
      city: r.City || '',
      country: r.Country || '',
      lastVisit: r.LastVisit || '',
      lastDuration: Number(r.LastDurationMin || r["Last Visited Duration (minutes)"] || 0) || 0,
      totalDuration: Number(r.TotalDurationMin || r["Total Visited Duration (minutes)"] || 0) || 0,
      totalVisits: Number(r.TotalVisitCount || r["Total Visit Count"] || 0) || 0,
      icon, status,
      page: r.LastPage || r["Last Visited Page / Section Name"] || ''
    };
  }

  function applyFilters(rows){
    const s = F.status.value;
    const q = (F.search.value||'').toLowerCase();
    const ctry = (F.country.value||'').toLowerCase();
    const cty = (F.city.value||'').toLowerCase();
    const minMin = Number(F.minmin.value||0) || 0;
    const range = F.range.value;

    const now = new Date();
    const cutoff = (rng=>{
      const map = { '24h':1, '3d':3, '7d':7, '30d':30, '90d':90, '180d':180, '365d':365 };
      const d = map[rng];
      if (!d) return null; const x = new Date(now.getTime() - d*86400000); return x;
    })(range);

    return rows.filter(r => {
      if (s !== 'all') {
        if (s === 'active' && !(r.status==='active'||r.status==='active-24h')) return false;
        else if (s !== 'active' && r.status !== s) return false;
      }
      if (ctry && !(r.country||'').toLowerCase().includes(ctry)) return false;
      if (cty && !(r.city||'').toLowerCase().includes(cty)) return false;
      if (minMin && (r.totalDuration||0) < minMin) return false;
      if (q) {
        const hay = (r.uid+' '+r.page+' '+r.ip).toLowerCase();
        if (!hay.includes(q)) return false;
      }
      if (cutoff) {
        const lv = r.lastVisit ? new Date(r.lastVisit) : null;
        if (!lv || lv < cutoff) return false;
      }
      return true;
    });
  }

  function renderTable(rows){
    const tb = byId('tbl').querySelector('tbody');
    tb.innerHTML = rows.slice(0, 500).map(r => `
      <tr>
        <td><span class="tag"><i>🆔</i>${r.uid}</span></td>
        <td>${r.timestamp||''}</td>
        <td>${r.ip||''}</td>
        <td>${r.city||''}</td>
        <td>${r.country||''}</td>
        <td>${r.lastVisit||''}</td>
        <td>${r.lastDuration||0}</td>
        <td>${r.totalDuration||0}</td>
        <td>${r.icon}</td>
      </tr>
    `).join('');
  }

  function setNum(elId, val){
    const el = byId(elId);
    const start = Number(el.textContent)||0; const end = val;
    const dur = 600; const t0 = performance.now();
    function step(t){
      const p = Math.min(1, (t-t0)/dur);
      const v = Math.round(start + (end - start) * p);
      el.textContent = String(v);
      if (p < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  }

  function spark(elId, data){
    const c = byId(elId); const ctx = c.getContext('2d');
    const w = c.width = c.clientWidth; const h = c.height = c.clientHeight;
    ctx.clearRect(0,0,w,h);
    const n = data.length || 1;
    const min = Math.min(...data, 0); const max = Math.max(...data, 1);
    ctx.strokeStyle = 'rgba(255,255,255,0.8)'; ctx.lineWidth = 2; ctx.beginPath();
    data.forEach((v,i)=>{
      const x = i/(n-1) * (w-8) + 4;
      const y = h - ((v-min)/(max-min+1e-6)) * (h-8) - 4;
      if (i===0) ctx.moveTo(x,y); else ctx.lineTo(x,y);
    });
    ctx.stroke();
  }

  function chartTotals(){
    const c = byId('chart_totals'); const ctx = c.getContext('2d');
    const w = c.width = c.clientWidth; const h = c.height = c.clientHeight;
    ctx.clearRect(0,0,w,h);
    const buckets = { '≤24h':0, '≤7d':0, '≤30d':0, '>30d':0 };
    const now = new Date();
    ROWS.forEach(r => {
      if (!r.lastVisit) { buckets['>30d']++; return; }
      const days = daysBetween(new Date(r.lastVisit), now);
      if (days <= 1) buckets['≤24h']++; else if (days <= 7) buckets['≤7d']++; else if (days <= 30) buckets['≤30d']++; else buckets['>30d']++;
    });
    const keys = Object.keys(buckets); const vals = keys.map(k=>buckets[k]);
    const maxv = Math.max(...vals, 1); const barW = Math.min(140, (w - 60) / keys.length - 20);
    keys.forEach((k, i) => {
      const v = vals[i]; const x = 40 + i*(barW+30); const y = h - 30; const bh = Math.max(6, (v/maxv) * (h-80));
      ctx.fillStyle = 'rgba(108,140,255,0.25)'; ctx.fillRect(x, y-bh, barW, bh);
      ctx.fillStyle = '#6c8cff'; ctx.fillRect(x, y-bh, barW, 6);
      ctx.fillStyle = 'rgba(255,255,255,0.85)'; ctx.font = '12px system-ui'; ctx.fillText(String(v), x + barW/2 - 4, y - bh - 8);
      ctx.fillStyle = 'rgba(255,255,255,0.65)'; ctx.fillText(k, x + Math.max(0, barW/2 - 18), y + 16);
    });
  }

  function recompute(){
    const rows = ROWS.map(normalize);
    const filtered = applyFilters(rows);

    // KPIs
    setNum('k_total', rows.length);
    const now = new Date();
    const active30 = rows.filter(r => r.status==='active' || r.status==='active-24h').length;
    const inactive = rows.filter(r => r.status==='inactive').length;
    const today = rows.filter(r => r.lastVisit && daysBetween(new Date(r.lastVisit), now) <= 1).length;
    const live = rows.filter(r => r.status==='live').length;
    const last24 = rows.filter(r => r.status==='active-24h' || r.status==='live').length;
    setNum('k_active', active30);
    setNum('k_inactive', inactive);
    setNum('k_today', today);
    setNum('k_live', live);
    setNum('k_24h', last24);

    renderTable(filtered);
    chartTotals();

    const totalVisitsAll = rows.reduce((a,r)=>a + (r.totalVisits||0), 0);
    const totalMinutesAll = rows.reduce((a,r)=>a + (r.totalDuration||0), 0);
    byId('totals_text').textContent = `All-time visits: ${totalVisitsAll} | Minutes: ${Math.round(totalMinutesAll)}`;

    // sparks
    spark('sp_total', [rows.length, Math.max(1, Math.round(rows.length*0.7)), Math.max(1, Math.round(rows.length*0.9)), rows.length]);
    spark('sp_active', [active30, Math.max(0, active30-2), active30+1]);
    spark('sp_inactive', [inactive, Math.max(0, inactive-1), inactive+2]);
    spark('sp_today', [today, Math.max(0, today-1), today+1]);
    spark('sp_live', [live, Math.max(0, live-1), live+1]);
    spark('sp_24h', [last24, Math.max(0, last24-2), last24+1]);
  }

  function bindFilters(){ Object.values(F).forEach(el => el && el.addEventListener('input', recompute)); }

  function fetchData(){
    // Prefer google.script.run when hosted in Apps Script
    if (typeof google !== 'undefined' && google.script && google.script.run) {
      google.script.run.withSuccessHandler((data)=>{ ROWS = (data && data.rows) || []; recompute(); }).getUsersData();
      return;
    }
    // Standalone fallback: try JSON endpoint on the same web app
    fetch(window.location.href.replace(/\?.*$/, '') + '?format=json')
      .then(r=>r.json())
      .then(j=>{ ROWS = j.rows||[]; recompute(); })
      .catch(()=>{ document.querySelector('#totals_text').textContent = 'Failed to load data'; });
  }

  window.addEventListener('resize', () => { recompute(); });

  bindFilters();
  fetchData();
})();
</script>
</body>
</html>`;
}
