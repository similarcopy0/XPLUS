/*
  Google Apps Script backend for ❌ˢʰᵒᵗᶻ (XShotz)
  - JSONP responses to bypass CORS for webview
  - Reads master content from Google Sheet: XShotz CONTENT SUMMARY
  - Writes interactions (likes, dislikes, favorites, shares, views) back to same sheet
  - Comments stored in separate sheet: XShotz Comments
  - Live stats endpoint

  Deploy as Web App (Anyone with link): doGet serves JSONP
*/

const SPREADSHEET_ID = '1PdAMnyrm6cIHUioi-MGcjMFgbmpn-Wa0zrXtTpuY7vE';
const SHEET_VIDEOS = 'XShotz CONTENT SUMMARY';
const SHEET_COMMENTS = 'XShotz Comments';

/** Utility: build JSONP response */
function jsonp(callback, obj) {
  const payload = callback + '(' + JSON.stringify(obj) + ')';
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/** Utility: normalize header to index map */
function headerIndexMap(headers) {
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim()] = i; });
  return map;
}

/** Utility: get sheet */
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name);
}

/** Utility: read all rows from sheet as objects */
function readSheetObjects(name) {
  const sh = getSheet(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const map = headerIndexMap(headers);
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const obj = {};
    Object.keys(map).forEach(k => obj[k] = row[map[k]]);
    rows.push(obj);
  }
  return rows;
}

/** Utility: write a single cell by header */
function writeCellByHeader(sheetName, rowIndex1, header, value) {
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const map = headerIndexMap(headers);
  const col = map[header];
  if (typeof col === 'undefined') throw new Error('Header not found: ' + header);
  sh.getRange(rowIndex1, col+1).setValue(value);
}

/** Utility: find row by id */
function findRowById(sheetName, id) {
  const sh = getSheet(sheetName);
  const range = sh.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const map = headerIndexMap(headers);
  const idCol = map['id'] != null ? map['id'] : map['ID'];
  if (idCol == null) return null;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]).trim() === String(id)) return { rowIndex1: r+1, map: map, row: values[r] };
  }
  return null;
}

/** GET handler */
function doGet(e) {
  const action = (e.parameter.action || '').trim();
  const callback = (e.parameter.callback || 'callback').trim();
  try {
    if (action === 'getVideos') return jsonp(callback, getVideos());
    if (action === 'getCounts') return jsonp(callback, getCounts(e.parameter.videoId));
    if (action === 'increment') return jsonp(callback, incrementMetric(e.parameter.videoId, e.parameter.metric, Number(e.parameter.amount||1)));
    if (action === 'liveStats') return jsonp(callback, liveStats());
    if (action === 'getComments') return jsonp(callback, getComments(e.parameter.videoId, (e.parameter.sort||'newest')));
    if (action === 'postComment') return jsonp(callback, postComment(e.parameter.videoId, e.parameter.user, e.parameter.text));
    if (action === 'search') return jsonp(callback, search(e.parameter.q||''));
    return jsonp(callback, { ok:false, error:'unknown_action' });
  } catch (err) {
    return jsonp(callback, { ok:false, error: String(err && err.message || err) });
  }
}

/** Actions */
function getVideos() {
  const rows = readSheetObjects(SHEET_VIDEOS);
  const data = rows.map(r => ({
    id: String(r.id || r.ID || '').trim(),
    title: r.title || '',
    description: r.description || '',
    category: r.category || '',
    videoUrl: r.videoUrl || r.video_url || '',
    likes: Number(r.likes || 0),
    dislikes: Number(r.dislikes || 0),
    commentsCount: Number(r.commentsCount || 0),
    favorites: Number(r.favorites || 0),
    shares: Number(r.shares || 0),
    saves: Number(r.saves || 0),
    is_trending: r.is_trending,
    new: r.new,
    views: Number(r.views || 0)
  }));
  return { ok:true, data };
}

function getCounts(videoId) {
  if (!videoId) return { ok:false, error:'missing_id' };
  const f = findRowById(SHEET_VIDEOS, videoId);
  if (!f) return { ok:false, error:'not_found' };
  const map = f.map, row = f.row;
  const out = {};
  ['likes','dislikes','commentsCount','favorites','shares','views'].forEach(k => {
    const idx = map[k]; if (typeof idx !== 'undefined') out[k] = Number(row[idx] || 0);
  });
  return { ok:true, data: out };
}

function incrementMetric(videoId, metric, amount) {
  if (!videoId || !metric) return { ok:false, error:'missing_params' };
  const f = findRowById(SHEET_VIDEOS, videoId);
  if (!f) return { ok:false, error:'not_found' };
  const map = f.map; const rowIndex1 = f.rowIndex1;
  if (typeof map[metric] === 'undefined') return { ok:false, error:'metric_not_found' };
  const sh = getSheet(SHEET_VIDEOS);
  const curr = Number(sh.getRange(rowIndex1, map[metric]+1).getValue() || 0);
  const next = curr + (amount||1);
  sh.getRange(rowIndex1, map[metric]+1).setValue(next);
  // If likes -> also increment commentsCount? No. If favorites -> nothing else.
  const data = getCounts(videoId).data;
  return { ok:true, data };
}

function liveStats() {
  const rows = readSheetObjects(SHEET_VIDEOS);
  let totalViews = 0; let totalVideos = 0;
  rows.forEach(r => { totalVideos++; totalViews += Number(r.views || 0); });
  return { ok:true, data: { totalViews, totalVideos } };
}

function getComments(videoId, sort) {
  if (!videoId) return { ok:false, error:'missing_id' };
  const rows = readSheetObjects(SHEET_COMMENTS);
  const items = rows.filter(r => String(r.videoId||'').trim() === String(videoId));
  items.forEach(i => { i.timestamp = i.timestamp || new Date(); });
  items.sort((a,b)=>{
    const ta = new Date(a.timestamp).getTime();
    const tb = new Date(b.timestamp).getTime();
    return sort === 'oldest' ? (ta - tb) : (tb - ta);
  });
  return { ok:true, data: items };
}

function postComment(videoId, user, text) {
  if (!videoId || !text) return { ok:false, error:'missing_params' };
  const sh = getSheet(SHEET_COMMENTS);
  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const map = headerIndexMap(headers);
  const row = [];
  // Build row in header order
  for (let c=0;c<headers.length;c++) {
    const h = headers[c];
    if (h==='commentId') row.push('c_' + Date.now() + '_' + Math.random().toString(36).slice(2,6));
    else if (h==='videoId') row.push(videoId);
    else if (h==='user') row.push(user || 'Anon');
    else if (h==='comments') row.push(text);
    else if (h==='timestamp') row.push(new Date());
    else row.push('');
  }
  sh.appendRow(row);

  // increment commentsCount on videos sheet
  const f = findRowById(SHEET_VIDEOS, videoId);
  if (f && typeof f.map['commentsCount'] !== 'undefined') {
    const curr = Number(getSheet(SHEET_VIDEOS).getRange(f.rowIndex1, f.map['commentsCount']+1).getValue()||0);
    getSheet(SHEET_VIDEOS).getRange(f.rowIndex1, f.map['commentsCount']+1).setValue(curr+1);
  }
  return { ok:true };
}

function search(q) {
  q = String(q||'').toLowerCase();
  const rows = readSheetObjects(SHEET_VIDEOS);
  const data = rows.filter(r => Object.keys(r).some(k => String(r[k]).toLowerCase().includes(q))).map(r => ({
    id: String(r.id || '').trim(),
    title: r.title || '',
    description: r.description || '',
    category: r.category || '',
    videoUrl: r.videoUrl || r.video_url || '',
    likes: Number(r.likes || 0),
    dislikes: Number(r.dislikes || 0),
    commentsCount: Number(r.commentsCount || 0),
    favorites: Number(r.favorites || 0),
    shares: Number(r.shares || 0),
    saves: Number(r.saves || 0),
    is_trending: r.is_trending,
    new: r.new,
    views: Number(r.views || 0)
  }));
  return { ok:true, data };
}
