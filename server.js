/**
 * Email Nudge — Outlook Add-in server
 *
 * Serves the taskpane HTML and static assets.
 * Must run over HTTPS for Office to load it — use a reverse proxy
 * (Nginx + Let's Encrypt) or deploy to a platform like Render / Railway
 * that provides HTTPS automatically.
 */

const express = require('express');
const path    = require('path');
const app     = express();
const PORT    = process.env.PORT || 3000;

// Security headers required by Office.js
app.use((req, res, next) => {
  // Office requires these for the add-in to load in the task pane
  res.setHeader('X-Content-Type-Options',   'nosniff');
  res.setHeader('X-Frame-Options',          'ALLOW-FROM https://outlook.office.com https://outlook.office365.com');
  res.setHeader('Content-Security-Policy',
    "default-src 'self' https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com; " +
    "script-src 'self' 'unsafe-inline' https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com; " +
    "style-src 'self' 'unsafe-inline'; " +
    "connect-src 'self' https://*.office.com https://*.microsoft.com https://graph.microsoft.com;"
  );
  next();
});

// Serve static files from /public
app.use(express.static(path.join(__dirname, 'public')));

// Serve manifest.xml at root for easy sideloading
app.get('/manifest.xml', (req, res) => {
  res.setHeader('Content-Type', 'application/xml');
  res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// Health check
app.get('/health', (req, res) => res.json({ status: 'ok', service: 'outlook-nudge' }));

app.listen(PORT, () => {
  console.log(`Email Nudge server running on port ${PORT}`);
  console.log(`Taskpane: http://localhost:${PORT}/taskpane.html`);
  console.log(`Manifest: http://localhost:${PORT}/manifest.xml`);
});
