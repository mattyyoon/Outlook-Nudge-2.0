const express = require('express');
const path    = require('path');
const app     = express();
const PORT    = process.env.PORT || 3000;

app.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  // ALLOWALL required — ALLOW-FROM is deprecated and ignored by modern browsers
  res.setHeader('X-Frame-Options', 'ALLOWALL');
  res.setHeader('Content-Security-Policy',
    "default-src 'self' https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com; " +
    "script-src 'self' 'unsafe-inline' https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com; " +
    "style-src 'self' 'unsafe-inline'; " +
    "img-src 'self' data:; " +
    "connect-src 'self' https://*.office.com https://*.microsoft.com https://graph.microsoft.com;"
  );
  next();
});

// Serve everything in /public (taskpane.html, commands.html, icons/)
app.use(express.static(path.join(__dirname, 'public')));

// Root route — status page so the bare URL works
app.get('/', (req, res) => {
  res.send(`<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>Email Nudge</title>
<style>
  body{font-family:system-ui,sans-serif;display:flex;align-items:center;
       justify-content:center;min-height:100vh;margin:0;background:#f8fafc}
  .box{text-align:center;padding:40px;max-width:420px}
  .icon{width:56px;height:56px;background:#0ea5e9;border-radius:14px;
        display:flex;align-items:center;justify-content:center;
        margin:0 auto 20px;font-size:28px}
  h1{font-size:22px;color:#0f172a;margin:0 0 8px}
  p{color:#64748b;margin:0 0 24px;line-height:1.6}
  .badge{display:inline-block;background:#dcfce7;color:#166534;
         font-size:13px;font-weight:600;padding:6px 16px;border-radius:20px}
  .links{margin-top:24px;display:flex;gap:12px;justify-content:center;flex-wrap:wrap}
  a{font-size:13px;color:#0ea5e9;text-decoration:none;
    padding:6px 14px;border:1px solid #bae6fd;border-radius:6px}
  a:hover{background:#f0f9ff}
</style></head>
<body>
<div class="box">
  <div class="icon">&#128276;</div>
  <h1>Email Nudge</h1>
  <p>Server is running. Install the add-in via the manifest link below.</p>
  <div class="badge">&#10003; Online</div>
  <div class="links">
    <a href="/manifest.xml">manifest.xml</a>
    <a href="/taskpane.html">taskpane.html</a>
    <a href="/health">health</a>
  </div>
</div>
</body></html>`);
});

// Manifest with correct content-type
app.get('/manifest.xml', (req, res) => {
  res.setHeader('Content-Type', 'application/xml');
  res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// Health check for Render
app.get('/health', (req, res) => {
  res.json({ status: 'ok', service: 'outlook-nudge', time: new Date().toISOString() });
});

app.listen(PORT, () => console.log(`Email Nudge running on port ${PORT}`));
