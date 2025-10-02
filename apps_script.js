// ❌ˢʰᵒᵗᶻ Apps Script for Google Sheets Integration
// This script handles data management for the video platform

// Configuration
const SHEET_ID = '1PdAMnyrm6cIHUioi-MGcjMFgbmpn-Wa0zrXtTpuY7vE';
const VIDEOS_SHEET_NAME = 'XShotz CONTENT SUMMARY';
const COMMENTS_SHEET_NAME = 'XShotz Comments';

// CORS and JSONP support function
function doGet(e) {
  const action = e.parameter.action;
  const videoId = e.parameter.videoId;
  const value = e.parameter.value;
  const timestamp = e.parameter.timestamp;
  const callback = e.parameter.callback; // For JSONP

  try {
    let result = {};

    switch (action) {
      case 'like':
        result = updateVideoStats(videoId, 'likes', 1);
        break;
      case 'dislike':
        result = updateVideoStats(videoId, 'dislikes', 1);
        break;
      case 'favorite':
        result = updateVideoStats(videoId, 'favorites', 1);
        break;
      case 'share':
        result = updateVideoStats(videoId, 'shares', 1);
        break;
      case 'view':
        result = updateVideoStats(videoId, 'views', 1);
        break;
      case 'comment':
        result = addComment(videoId, value);
        break;
      case 'getVideos':
        result = getVideos();
        break;
      case 'getComments':
        result = getComments(videoId);
        break;
      default:
        result = { error: 'Invalid action' };
    }

    // Return JSONP if callback is provided
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      // Return CORS-enabled JSON
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader('Access-Control-Allow-Origin', '*')
        .setHeader('Access-Control-Allow-Methods', 'GET')
        .setHeader('Access-Control-Allow-Headers', 'Content-Type');
    }

  } catch (error) {
    const errorResult = { error: error.message };

    if (callback) {
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(errorResult) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      return ContentService.createTextOutput(JSON.stringify(errorResult))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
}

// Update video statistics in the sheet
function updateVideoStats(videoId, statType, increment) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(VIDEOS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    // Find the video row
    let videoRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === videoId) { // ID is in column A
        videoRow = i + 1; // +1 because array is 0-based, sheet is 1-based
        break;
      }
    }

    if (videoRow === -1) {
      return { error: 'Video not found' };
    }

    // Get column index for the stat type
    const headers = data[0];
    const statColumnIndex = headers.indexOf(statType);

    if (statColumnIndex === -1) {
      return { error: 'Invalid stat type' };
    }

    // Get current value and increment
    const currentValue = parseInt(sheet.getRange(videoRow, statColumnIndex + 1).getValue()) || 0;
    const newValue = currentValue + increment;

    // Update the sheet
    sheet.getRange(videoRow, statColumnIndex + 1).setValue(newValue);

    // Return success with updated value
    return {
      success: true,
      videoId: videoId,
      statType: statType,
      newValue: newValue,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    return { error: error.message };
  }
}

// Add a comment to the comments sheet
function addComment(videoId, commentText) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(COMMENTS_SHEET_NAME);

    // Generate unique comment ID
    const commentId = 'comment_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);

    // Get user ID (in a real app, this would come from authentication)
    const userId = 'user_' + (Math.random().toString(36).substr(2, 9));

    const timestamp = new Date().toISOString();

    // Add row to comments sheet
    sheet.appendRow([commentId, videoId, userId, commentText, timestamp]);

    return {
      success: true,
      commentId: commentId,
      videoId: videoId,
      userId: userId,
      comment: commentText,
      timestamp: timestamp
    };

  } catch (error) {
    return { error: error.message };
  }
}

// Get all videos data
function getVideos() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(VIDEOS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    const headers = data[0];
    const videos = [];

    for (let i = 1; i < data.length; i++) {
      const video = {};
      headers.forEach((header, index) => {
        video[header.toLowerCase().trim()] = data[i][index] || '';
      });

      // Only include videos with required fields
      if (video.id && video.videourl) {
        videos.push(video);
      }
    }

    return {
      success: true,
      videos: videos,
      count: videos.length,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    return { error: error.message };
  }
}

// Get comments for a specific video
function getComments(videoId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(COMMENTS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    const headers = data[0];
    const comments = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === videoId) { // videoId is in column B
        const comment = {
          commentId: data[i][0],
          videoId: data[i][1],
          user: data[i][2],
          comment: data[i][3],
          timestamp: data[i][4]
        };
        comments.push(comment);
      }
    }

    // Sort comments by timestamp (newest first)
    comments.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    return {
      success: true,
      comments: comments,
      count: comments.length,
      videoId: videoId,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    return { error: error.message };
  }
}

// Utility function to get sheet data as CSV (for the HTML/JS to parse)
function getSheetAsCSV(sheetName) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
    const range = sheet.getDataRange();
    const data = range.getValues();

    // Convert to CSV format
    let csv = '';
    for (let i = 0; i < data.length; i++) {
      const row = data[i].map(field => {
        // Escape quotes and wrap in quotes if necessary
        const stringField = String(field || '');
        if (stringField.includes(',') || stringField.includes('"') || stringField.includes('\n')) {
          return '"' + stringField.replace(/"/g, '""') + '"';
        }
        return stringField;
      });
      csv += row.join(',') + '\n';
    }

    return csv;

  } catch (error) {
    return 'Error: ' + error.message;
  }
}

// Web app endpoint for CSV data (for backward compatibility)
function serveCSV() {
  const htmlOutput = HtmlService.createHtmlOutput('<script>window.parent.postMessage("' +
    getSheetAsCSV(VIDEOS_SHEET_NAME).replace(/\n/g, '\\n').replace(/"/g, '\\"') +
    '", "*");</script>');
  return htmlOutput.setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// Setup function to run when the script is first deployed
function onInstall() {
  // Set up trigger to refresh data periodically (optional)
  ScriptApp.newTrigger('refreshData')
    .timeBased()
    .everyHours(1)
    .create();

  // Log setup completion
  Logger.log('❌ˢʰᵒᵗᶻ Apps Script setup completed');
}

// Refresh data function (called by trigger)
function refreshData() {
  // This function can be used to periodically update aggregated data
  // For now, it just logs that it's running
  Logger.log('Data refresh triggered at: ' + new Date().toISOString());
}

// Test function to verify the script is working
function testScript() {
  try {
    const testVideoId = 'test_video_1';

    // Test updating likes
    const likeResult = updateVideoStats(testVideoId, 'likes', 1);
    Logger.log('Like test result: ' + JSON.stringify(likeResult));

    // Test adding comment
    const commentResult = addComment(testVideoId, 'Test comment from Apps Script');
    Logger.log('Comment test result: ' + JSON.stringify(commentResult));

    return {
      success: true,
      message: 'Script test completed successfully',
      results: {
        like: likeResult,
        comment: commentResult
      }
    };

  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// Additional utility functions for data management

// Get video statistics summary
function getVideoStatsSummary() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(VIDEOS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    let totalViews = 0;
    let totalLikes = 0;
    let totalComments = 0;
    let totalVideos = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // If ID exists
        totalVideos++;
        totalViews += parseInt(data[i][11]) || 0; // views column
        totalLikes += parseInt(data[i][5]) || 0;  // likes column
        totalComments += parseInt(data[i][8]) || 0; // commentsCount column
      }
    }

    return {
      success: true,
      summary: {
        totalVideos: totalVideos,
        totalViews: totalViews,
        totalLikes: totalLikes,
        totalComments: totalComments
      },
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    return { error: error.message };
  }
}

// Clean up old comments (optional maintenance function)
function cleanupOldComments(daysToKeep = 30) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(COMMENTS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {
      const commentDate = new Date(data[i][4]); // timestamp column
      if (commentDate < cutoffDate) {
        rowsToDelete.push(i + 1);
      }
    }

    // Delete old comments (from bottom to top to maintain row numbers)
    rowsToDelete.reverse().forEach(row => {
      sheet.deleteRow(row);
    });

    return {
      success: true,
      deletedRows: rowsToDelete.length,
      message: `Cleaned up ${rowsToDelete.length} old comments`
    };

  } catch (error) {
    return { error: error.message };
  }
}

// Export data for backup
function exportData() {
  try {
    const videosData = getSheetAsCSV(VIDEOS_SHEET_NAME);
    const commentsData = getSheetAsCSV(COMMENTS_SHEET_NAME);

    return {
      success: true,
      exportData: {
        videos: videosData,
        comments: commentsData,
        exportDate: new Date().toISOString()
      }
    };

  } catch (error) {
    return { error: error.message };
  }
}

// Import data from backup (for data restoration)
function importData(videosCSV, commentsCSV) {
  try {
    // This would parse and import CSV data back into sheets
    // Implementation depends on specific requirements

    return {
      success: true,
      message: 'Import functionality available - contact developer for implementation'
    };

  } catch (error) {
    return { error: error.message };
  }
}