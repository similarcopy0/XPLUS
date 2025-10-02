// Google Apps Script for ❌ˢʰᵒᵗᶻ Video App
// Deploy as Web App and use the URL in the HTML file

// Spreadsheet IDs
const CONTENT_SPREADSHEET_ID = "1PdAMnyrm6cIHUioi-MGcjMFgbmpn-Wa0zrXtTpuY7vE";
const COMMENTS_SPREADSHEET_ID = "1PdAMnyrm6cIHUioi-MGcjMFgbmpn-Wa0zrXtTpuY7vE"; // Same spreadsheet for now

// Sheet names
const CONTENT_SHEET_NAME = "XShotz CONTENT SUMMARY";
const COMMENTS_SHEET_NAME = "XShotz Comments";

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback || "callback";

  try {
    if (action === "content") {
      // Return video content data
      const data = getContentData();
      return createJsonpResponse(callback, data);

    } else if (action === "comments") {
      // Return comments data
      const data = getCommentsData();
      return createJsonpResponse(callback, data);

    } else if (action === "like") {
      // Handle like action
      const videoId = e.parameter.videoId;
      const result = handleLike(videoId);
      return createJsonpResponse(callback, result);

    } else if (action === "comment") {
      // Handle comment action
      const videoId = e.parameter.videoId;
      const user = e.parameter.user || "Anonymous";
      const comment = e.parameter.comment;
      const result = handleComment(videoId, user, comment);
      return createJsonpResponse(callback, result);

    } else {
      // Default: return content data
      const data = getContentData();
      return createJsonpResponse(callback, data);
    }

  } catch (error) {
    console.error("Error in doGet:", error);
    return createJsonpResponse(callback, { error: error.message });
  }
}

function getContentData() {
  const sheet = SpreadsheetApp.openById(CONTENT_SPREADSHEET_ID).getSheetByName(CONTENT_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  // Add headers for clarity
  const headers = data[0];
  const rows = data.slice(1);

  return [headers, ...rows];
}

function getCommentsData() {
  const sheet = SpreadsheetApp.openById(COMMENTS_SPREADSHEET_ID).getSheetByName(COMMENTS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  // Add headers for clarity
  const headers = data[0];
  const rows = data.slice(1);

  return [headers, ...rows];
}

function handleLike(videoId) {
  try {
    const sheet = SpreadsheetApp.openById(CONTENT_SPREADSHEET_ID).getSheetByName(CONTENT_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Find video by ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === videoId) { // ID column
        const currentLikes = parseInt(data[i][5]) || 0; // likes column
        sheet.getRange(i + 1, 6).setValue(currentLikes + 1); // Increment likes
        break;
      }
    }

    // Return updated content data
    return getContentData();

  } catch (error) {
    console.error("Error handling like:", error);
    return { error: error.message };
  }
}

function handleComment(videoId, user, comment) {
  try {
    const sheet = SpreadsheetApp.openById(COMMENTS_SPREADSHEET_ID).getSheetByName(COMMENTS_SHEET_NAME);

    // Generate unique comment ID
    const timestamp = new Date().toISOString();
    const commentId = "comment_" + Date.now();

    // Append new comment
    sheet.appendRow([commentId, videoId, user, comment, timestamp]);

    // Return updated comments data
    return getCommentsData();

  } catch (error) {
    console.error("Error handling comment:", error);
    return { error: error.message };
  }
}

function createJsonpResponse(callback, data) {
  const output = JSON.stringify(data);
  const response = ContentService.createTextOutput(callback + "(" + output + ")")
                                 .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return response;
}

// Additional utility functions

function shuffleArray(array) {
  const shuffled = [...array];
  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }
  return shuffled;
}

function getRandomizedContent() {
  const data = getContentData();
  const headers = data[0];
  const rows = shuffleArray(data.slice(1));

  return [headers, ...rows];
}

function searchContent(query) {
  const data = getContentData();
  const headers = data[0];
  const rows = data.slice(1);

  const filteredRows = rows.filter(row => {
    return row.some(cell =>
      cell && cell.toString().toLowerCase().includes(query.toLowerCase())
    );
  });

  return [headers, ...filteredRows];
}

// Alternative endpoints for different data views
function getTrendingVideos() {
  const data = getContentData();
  const headers = data[0];
  const rows = data.slice(1);

  const trendingRows = rows.filter(row => {
    const isTrending = row[10]; // is_trending column
    return isTrending === true || isTrending === "TRUE" || isTrending === "true";
  });

  return [headers, ...shuffleArray(trendingRows)];
}

function getNewVideos() {
  const data = getContentData();
  const headers = data[0];
  const rows = data.slice(1);

  const newRows = rows.filter(row => {
    const isNew = row[11]; // new column
    return isNew === true || isNew === "TRUE" || isNew === "true";
  });

  return [headers, ...shuffleArray(newRows)];
}

function getVideosByCategory(category) {
  const data = getContentData();
  const headers = data[0];
  const rows = data.slice(1);

  const categoryRows = rows.filter(row => {
    return row[3] && row[3].toString().toLowerCase() === category.toLowerCase();
  });

  return [headers, ...shuffleArray(categoryRows)];
}