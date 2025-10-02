# ❌ˢʰᵒᵗᶻ - Short Video Platform

A modern, TikTok/Reels-style short-form video platform built with HTML, CSS, and JavaScript, integrated with Google Sheets and Apps Script for data management.

## Features

### 🎥 Advanced Video Player
- **Auto-play**: Videos automatically play when 80% visible in viewport
- **Loop**: Seamless video looping for continuous playback
- **Fullscreen**: Tap to toggle immersive fullscreen mode
- **Touch Controls**: Double-tap to like, single-tap to pause/play
- **Progress Bar**: Beautiful gradient progress indicator
- **Preloading**: Smart preloading of next videos for smooth transitions
- **Error Handling**: Automatic retry and graceful fallbacks for failed videos

### 📱 Modern User Interface
- **TikTok/Reels Design**: Authentic short-video app interface
- **Responsive**: Mobile-first design optimized for all screen sizes
- **Dark Theme**: Eye-friendly dark mode with proper contrast
- **Smooth Animations**: 60fps animations with CSS transitions and JavaScript
- **Glass Morphism**: Modern backdrop blur effects
- **Safe Areas**: Support for devices with notches and rounded corners

### 🔄 Robust Data Integration
- **Google Sheets**: Complete data management in Google Sheets
- **Apps Script**: Powerful backend API with 15+ utility functions
- **JSONP Support**: CORS workaround for seamless data fetching
- **Real-time Updates**: Live statistics updates every 30 seconds
- **Offline Support**: Graceful degradation when offline
- **Error Recovery**: Automatic retry mechanisms for failed requests

### 🎯 Interactive Features
- **Like/Dislike**: Animated engagement buttons with visual feedback
- **Comments**: Full comment system with filtering (newest/oldest)
- **Share**: Native Web Share API with clipboard fallback
- **Favorites**: Local storage-based save system
- **Search**: Live search with instant results
- **Double-tap Hearts**: TikTok-style like animation on double-tap

### 📊 Analytics & Performance
- **Live Stats Bar**: Real-time view counts and video statistics
- **Engagement Tracking**: Comprehensive interaction analytics
- **Performance Monitoring**: FPS monitoring and memory optimization
- **Battery Optimization**: Reduced quality on low battery (mobile)
- **Network Monitoring**: Offline detection and retry mechanisms

### ⌨️ Enhanced Controls
- **Keyboard Shortcuts**: Arrow keys, spacebar, F/F11 for fullscreen
- **Touch Gestures**: Swipe navigation, pinch-to-zoom support
- **Intersection Observer**: Smart video play/pause based on visibility
- **Memory Management**: Automatic cleanup of unused video elements

## File Structure

```
/workspace/
├── xshotz_app.html      # Main application file (HTML + CSS + JS)
├── apps_script.js       # Google Apps Script backend
└── README.md           # This documentation file
```

## Setup Instructions

### 1. Google Sheets Setup

1. **Create Google Sheet**:
   - Go to [Google Sheets](https://docs.google.com/spreadsheets)
   - Create a new spreadsheet
   - Use the provided Sheet ID: `1PdAMnyrm6cIHUioi-MGcjMFgbmpn-Wa0zrXtTpuY7vE`

2. **Sheet Structure**:
   - **Videos Sheet**: Name it `XShotz CONTENT SUMMARY`
   - **Comments Sheet**: Name it `XShotz Comments`

3. **Required Columns for Videos Sheet**:
   ```
   id, title, description, category, videoUrl, likes, dislikes,
   commentsCount, favorites, shares, saves, is_trending, new, views
   ```

4. **Required Columns for Comments Sheet**:
   ```
   commentId, videoId, user, comments, timestamp
   ```

### 2. Apps Script Setup

1. **Open Apps Script**:
   - In your Google Sheet, go to `Extensions > Apps Script`
   - Replace the default code with the contents of `apps_script.js`

2. **Deploy Web App**:
   - Click `Deploy > New deployment`
   - Select `Web app` as deployment type
   - Set `Execute as: Me` and `Who has access: Anyone`
   - Copy the Web App URL for use in the HTML file

3. **Update URLs in HTML**:
   - Open `xshotz_app.html`
   - Update the `APPS_SCRIPT_URL` variable with your deployed web app URL

### 3. Video Data Population

1. **Add Video Data**:
   - In the Videos sheet, add your video information
   - Required fields: `id`, `videoUrl`, `title`
   - Optional: `description`, `category`, `thumbnail_link`

2. **Video URL Format**:
   - Use direct video file URLs (MP4 format recommended)
   - For GitHub: Use raw file URLs or GitHub Releases URLs

### 4. Testing

1. **Open Application**:
   - Open `xshotz_app.html` in a web browser
   - Or host it on a web server for better performance

2. **Test Features**:
   - Video autoplay and navigation
   - Like/dislike functionality
   - Comment system
   - Search functionality
   - Grid view (Explore tab)

## Usage Guide

### Basic Navigation
- **Browse Videos**: Scroll vertically to navigate between videos
- **Like Video**: Double-tap video or tap heart icon
- **View Comments**: Tap comment icon to open comment drawer
- **Search**: Tap search icon and enter keywords
- **Explore**: Tap "Explore" tab for grid view

### Keyboard Shortcuts
- **Arrow Down/Space**: Next video
- **Arrow Up**: Previous video
- **F/F11**: Toggle fullscreen
- **Escape**: Exit fullscreen

### Mobile Gestures
- **Scroll**: Navigate between videos
- **Double Tap**: Like video (shows heart animation)
- **Single Tap**: Pause/play video
- **Pinch**: Zoom (if supported by device)

## Configuration

### Customization Options

#### Video Settings
```javascript
// Maximum video duration (seconds)
const MAX_VIDEO_DURATION = 180;

// Auto-advance after N videos
const AUTO_ADVANCE_COUNT = 7;

// Stats update interval (milliseconds)
const STATS_UPDATE_INTERVAL = 30000;
```

#### UI Customization
```css
/* Modify colors in the CSS section */
:root {
  --primary-color: #ff6b6b;
  --accent-color: #667eea;
  --background: #000;
}
```

## Browser Compatibility

### Supported Browsers
- **Chrome/Edge**: Full support (recommended)
- **Safari**: Full support with iOS 12+
- **Firefox**: Full support
- **Mobile browsers**: Optimized for mobile experience

### Required Features
- **ES6+ JavaScript**: Arrow functions, async/await, etc.
- **CSS Grid & Flexbox**: For responsive layouts
- **Video Element**: HTML5 video with autoplay
- **Local Storage**: For favorites and history

## Performance Optimization

### Best Practices
1. **Video Preloading**: Next videos are preloaded for smooth transitions
2. **Lazy Loading**: Videos load only when needed
3. **Memory Management**: Proper cleanup of video elements
4. **Error Handling**: Graceful fallbacks for network issues

### Performance Tips
- Use videos under 50MB for better loading times
- Optimize thumbnails for faster grid loading
- Enable gzip compression on your web server
- Use CDN for video hosting when possible

## Troubleshooting

### Common Issues

#### Videos Not Loading
- Check video URLs are accessible and in correct format
- Verify Google Sheets is published and accessible
- Check browser console for CORS errors

#### Apps Script Not Working
- Ensure web app is deployed correctly
- Check execution permissions are set to "Anyone"
- Verify script has access to the Google Sheet

#### Mobile Issues
- Enable autoplay in browser settings if needed
- Check if videos are in supported format (MP4)
- Ensure device supports HTML5 video

### Debug Mode
Add to URL: `?debug=true` for additional console logging

## Development

### Adding New Features
1. **UI Changes**: Modify HTML/CSS in `xshotz_app.html`
2. **Functionality**: Add JavaScript functions
3. **Backend**: Update Apps Script for new data operations

### Testing
- Test on multiple devices and browsers
- Validate video URLs and data integrity
- Check performance with browser dev tools

## Security Considerations

### Data Protection
- Google Sheets data is protected by Google's security
- Apps Script runs in Google's secure environment
- No sensitive data stored in client-side code

### Privacy
- User interactions tracked only for analytics
- No personal data collection without consent
- Local storage used only for preferences

## License

This project is open source and available under the MIT License.

## Support

For issues or questions:
1. Check the troubleshooting section
2. Verify all setup steps are completed
3. Test with sample data first
4. Check browser console for error messages

---

**Built with ❤️ for the modern web**