/**
 * XShotz Apps Script JSONP Backend
 * Sheets:
 *  - Content: XShotz CONTENT SUMMARY
 *  - Comments: XShotz Comments
 */

const SPREADSHEET_ID = '1PdAMnyrm6cIHUioi-MGcjMFgbmpn-Wa0zrXtTpuY7vE';
const CONTENT_SHEET = 'XShotz CONTENT SUMMARY';
const COMMENTS_SHEET = 'XShotz Comments';

function doGet(e) {
  const route = (e.parameter.route || 'content').toLowerCase();
  const callback = e.parameter.callback || 'callback';
  let payload;
  try {
    if (route === 'content') payload = routeContent();
    else if (route === 'counts') payload = routeCounts(e);
    else if (route === 'update') payload = routeUpdate(e);
    else if (route === 'comments') payload = routeComments(e);
    else if (route === 'addcomment') payload = routeAddComment(e);
    else payload = { error: 'unknown_route' };
  } catch (err) {
    payload = { error: String(err) };
  }
  const output = Utilities.jsonStringify(payload);
  return ContentService.createTextOutput(`${callback}(${output})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function routeContent() {
  const sheet = getSheet(CONTENT_SHEET);
  const values = sheet.getDataRange().getValues();
  // Return raw rows (header + rows). Frontend can normalize
  return values;
}

function findColumnIndexes_(header) {
  const index = {};
  header.forEach((h, i) => index[String(h).trim()] = i);
  return index;
}

function routeCounts(e) {
  const ids = String(e.parameter.ids || '').split(',').filter(Boolean);
  if (!ids.length) return {};
  const sheet = getSheet(CONTENT_SHEET);
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const idx = findColumnIndexes_(header);
  const out = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = String(row[idx['id']]);
    if (ids.indexOf(id) > -1) {
      out[id] = {
        likes: Number(row[idx['likes']]||0),
        dislikes: Number(row[idx['dislikes']]||0),
        commentsCount: Number(row[idx['commentsCount']]||0),
        favorites: Number(row[idx['favorites']]||0),
        shares: Number(row[idx['shares']]||0),
        views: Number(row[idx['views']]||0),
      };
    }
  }
  return out;
}

function routeUpdate(e) {
  const videoId = String(e.parameter.videoId || '');
  const field = String(e.parameter.field || '').trim();
  const delta = Number(e.parameter.delta || 1);
  if (!videoId || !field) return { ok: false, error: 'missing_params' };
  const sheet = getSheet(CONTENT_SHEET);
  const range = sheet.getDataRange();
  const values = range.getValues();
  const header = values[0];
  const idx = findColumnIndexes_(header);
  const colIndex = idx[field];
  if (colIndex == null) return { ok: false, error: 'field_not_found' };
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idx['id']]) === videoId) {
      const current = Number(values[r][colIndex] || 0);
      values[r][colIndex] = current + delta;
      range.offset(0,0,values.length, header.length).setValues(values);
      return { ok: true, id: videoId, field, value: values[r][colIndex] };
    }
  }
  return { ok: false, error: 'id_not_found' };
}

function routeComments(e) {
  const videoId = String(e.parameter.videoId || '');
  if (!videoId) return [];
  const sheet = getSheet(COMMENTS_SHEET);
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const idx = findColumnIndexes_(header);
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[idx['videoId']]) === videoId) {
      out.push({
        commentId: row[idx['commentId']],
        videoId: row[idx['videoId']],
        user: row[idx['user']],
        comments: row[idx['comments']],
        timestamp: new Date(row[idx['timestamp']]).getTime(),
      });
    }
  }
  // newest first
  out.sort((a,b) => b.timestamp - a.timestamp);
  return out;
}

function routeAddComment(e) {
  const videoId = String(e.parameter.videoId || '').trim();
  const user = String(e.parameter.user || '').trim();
  const comments = String(e.parameter.comments || '').trim();
  if (!videoId || !comments) return { ok: false, error: 'missing_params' };
  const sheet = getSheet(COMMENTS_SHEET);
  const id = 'c_' + Math.random().toString(36).slice(2) + Date.now().toString(36);
  const now = new Date();
  sheet.appendRow([id, videoId, user || 'anon', comments, now]);

  // Increment commentsCount in content sheet
  const cs = getSheet(CONTENT_SHEET);
  const range = cs.getDataRange();
  const values = range.getValues();
  const header = values[0];
  const idx = findColumnIndexes_(header);
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idx['id']]) === videoId) {
      const col = idx['commentsCount'];
      if (col != null) {
        const current = Number(values[r][col] || 0);
        values[r][col] = current + 1;
        range.offset(0,0,values.length, header.length).setValues(values);
      }
      break;
    }
  }

  return { ok: true, commentId: id };
}
