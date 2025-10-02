# ❌ˢʰᵒᵗᶻ - Short Form Video WebView Android App

A modern short-form video application similar to TikTok/Reels/Shorts, built as a single HTML file with embedded CSS and JavaScript for Android WebView deployment.

## 🚀 Features

- **Vertical Video Feed**: TikTok-style swipe navigation
- **Auto-play Videos**: Seamless video playback with mute autoplay
- **Interactive UI**: Like, dislike, comment, share, and favorite functionality
- **Real-time Stats**: Live view counts and engagement metrics
- **Search Functionality**: Search across titles, descriptions, and categories
- **Grid Layouts**: 2-column grid for Explore and Profile sections
- **Comment System**: Full comment drawer with real-time updates
- **Responsive Design**: Mobile-first design optimized for all screen sizes
- **Google Sheets Integration**: Data sourced from Google Sheets via Apps Script

## 📋 Prerequisites

1. **Google Account** with Google Sheets and Apps Script access
2. **Android Device/Emulator** for WebView testing
3. **Web Server** or **GitHub Pages** for hosting (optional for WebView)

## 🛠️ Setup Instructions

### Step 1: Setup Google Sheets

1. **Create Spreadsheet**: Go to [Google Sheets](https://docs.google.com/spreadsheets) and create a new spreadsheet
2. **Sheet Structure**:

   **Sheet 1: "XShotz CONTENT SUMMARY"**
   ```csv
   id,title,description,category,videoUrl,likes,dislikes,commentsCount,favorites,shares,saves,is_trending,new,views
   video1,Sample Video 1,Description 1,Fashion,https://example.com/video1.mp4,100,10,50,25,5,1,false,false,1000
   video2,Sample Video 2,Description 2,Tech,https://example.com/video2.mp4,200,20,75,40,8,1,false,false,2500
   ```

   **Sheet 2: "XShotz Comments"**
   ```csv
   commentId,videoId,user,comments,timestamp
   comment1,video1,User1,Great video!,2024-01-01T10:00:00Z
   comment2,video1,User2,Love this!,2024-01-01T11:00:00Z
   ```

3. **Make Public**: Share the spreadsheet publicly (Anyone with link can view)

### Step 2: Deploy Apps Script

1. **Open Apps Script**: In your Google Sheet, go to **Extensions > Apps Script**
2. **Replace Code**: Replace the default code with the contents of `apps-script.js`
3. **Deploy Web App**:
   - Click **Deploy > New deployment**
   - Select **Web app**
   - Set **Execute as: Me**
   - Set **Who has access: Anyone**
   - Copy the **Web app URL**

### Step 3: Update HTML File

1. **Edit URLs**: In `index.html`, update these lines:
   ```javascript
   const SHEET_URL = 'YOUR_DEPLOYED_APPS_SCRIPT_URL';
   const COMMENTS_SHEET_URL = 'YOUR_DEPLOYED_APPS_SCRIPT_URL';
   ```

2. **Add Video URLs**: Ensure your video URLs are accessible (GitHub raw URLs work well)

### Step 4: Test the Application

1. **Local Testing**: Open `index.html` in a modern browser
2. **WebView Testing**: Load the HTML file in Android WebView
3. **Mobile Testing**: Test on actual mobile devices

## 📱 WebView Integration

For Android WebView integration:

```java
WebView webView = findViewById(R.id.webview);
WebSettings webSettings = webView.getSettings();

// Enable JavaScript
webSettings.setJavaScriptEnabled(true);

// Allow media playback without user gesture
webSettings.setMediaPlaybackRequiresUserGesture(false);

// Load your HTML file
webView.loadUrl("file:///android_asset/index.html");
```

## 🎨 Customization

### Styling
- Modify CSS in the `<style>` section of `index.html`
- Colors, fonts, and animations can be customized
- Mobile-responsive breakpoints are included

### Functionality
- Add new interaction types in the action bar
- Modify video loading logic in JavaScript
- Customize data columns in Google Sheets

## 🔧 Troubleshooting

### Common Issues

1. **Videos not loading**: Check video URLs are accessible and in correct format
2. **CORS errors**: Ensure Apps Script is deployed correctly and URLs are updated
3. **Autoplay not working**: Enable autoplay in browser settings for testing

### Performance Optimization

- Videos are preloaded for smooth scrolling
- Lazy loading implemented for grid layouts
- Real-time updates occur every 5 seconds
- Optimistic UI updates for better user experience

## 📊 Data Structure

### Video Data Columns
- `id`: Unique video identifier
- `title`: Video title
- `description`: Video description
- `category`: Video category
- `videoUrl`: Direct video URL (MP4 format)
- `likes`, `dislikes`, `commentsCount`, `favorites`, `shares`, `saves`: Engagement metrics
- `is_trending`, `new`: Boolean flags
- `views`: View count

### Comments Data Columns
- `commentId`: Unique comment identifier
- `videoId`: Reference to video ID
- `user`: Comment author
- `comments`: Comment text
- `timestamp`: Comment timestamp

## 🚀 Deployment

### Option 1: GitHub Pages (Recommended)
1. Upload `index.html` to a GitHub repository
2. Enable GitHub Pages in repository settings
3. Update URLs in HTML file to point to deployed Apps Script

### Option 2: Web Server
1. Upload `index.html` to any web server
2. Ensure HTTPS for video loading (recommended)

### Option 3: Local WebView
1. Place `index.html` in Android assets folder
2. Load via WebView for offline functionality

## 🔒 Security Notes

- Google Sheets should be shared appropriately
- Apps Script web apps are secure when deployed correctly
- Video URLs should be from trusted sources
- Consider user data privacy in production

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues and enhancement requests.

---

**Built with ❤️ for the modern web**