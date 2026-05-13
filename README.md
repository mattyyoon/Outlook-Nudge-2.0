# Email Nudge — Outlook Add-in

Surfaces unanswered emails right inside Outlook's task pane. Shows how long ago each thread was sent, flags overdue replies, and lets you reply without leaving the panel.

---

## How it works

The add-in is a small web app that Office.js loads in Outlook's side panel. It:

1. Requests a short-lived OAuth token from `Office.context.mailbox`
2. Uses that token to call the **Outlook REST API** (v2.0) to fetch your inbox and sent items
3. Finds emails older than your threshold with no reply
4. Surfaces them as nudge cards grouped by urgency
5. Sends replies directly through the REST API when you hit "Reply now"

All data stays between your browser and Microsoft's servers — nothing passes through your hosting server.

---

## File structure

```
outlook-nudge/
├── public/
│   ├── taskpane.html     ← the entire add-in UI + Office.js logic
│   ├── commands.html     ← required function file (empty)
│   └── icons/
│       ├── icon-16.png
│       ├── icon-32.png
│       ├── icon-64.png
│       └── icon-80.png
├── manifest.xml          ← registers the add-in with Outlook
├── server.js             ← minimal Express server
├── package.json
└── make_icons.py         ← generates PNG icons (run once)
```

---

## Option A — Deploy to Render (free, HTTPS automatic)

> Render gives you a free HTTPS URL with zero config. Easiest path.

**1. Push to GitHub**

```bash
cd outlook-nudge
git init
git add .
git commit -m "initial"
# create a repo at github.com, then:
git remote add origin https://github.com/YOUR-USERNAME/outlook-nudge.git
git push -u origin main
```

**2. Deploy on Render**

1. Go to [render.com](https://render.com) → New → Web Service
2. Connect your GitHub repo
3. Settings:
   - **Environment:** Node
   - **Build command:** `npm install`
   - **Start command:** `node server.js`
   - **Instance type:** Free
4. Click **Deploy** — Render gives you a URL like `https://outlook-nudge-xxxx.onrender.com`

**3. Update manifest.xml**

Replace every `https://YOUR-DOMAIN` with your Render URL:

```bash
# Mac/Linux:
sed -i '' 's|https://YOUR-DOMAIN|https://outlook-nudge-xxxx.onrender.com|g' manifest.xml

# Windows PowerShell:
(Get-Content manifest.xml) -replace 'https://YOUR-DOMAIN','https://outlook-nudge-xxxx.onrender.com' | Set-Content manifest.xml
```

---

## Option B — Deploy to Railway

```bash
npm install -g @railway/cli
railway login
railway init
railway up
# Railway gives you an HTTPS URL automatically
```

Then replace `YOUR-DOMAIN` in `manifest.xml` with the Railway URL.

---

## Option C — Self-host with Nginx + Let's Encrypt

> For your own server (VPS, home lab, etc.)

**Install dependencies:**
```bash
sudo apt install nodejs npm nginx certbot python3-certbot-nginx
```

**Install app:**
```bash
git clone https://github.com/YOUR-USERNAME/outlook-nudge.git /var/www/nudge
cd /var/www/nudge
npm install
```

**Nginx config** (`/etc/nginx/sites-available/nudge`):
```nginx
server {
    server_name nudge.yourdomain.com;
    location / {
        proxy_pass http://localhost:3000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
    }
}
```

**Enable + SSL:**
```bash
sudo ln -s /etc/nginx/sites-available/nudge /etc/nginx/sites-enabled/
sudo certbot --nginx -d nudge.yourdomain.com
sudo nginx -t && sudo systemctl reload nginx
```

**Run with PM2 (auto-restart):**
```bash
sudo npm install -g pm2
pm2 start server.js --name nudge
pm2 save
pm2 startup
```

---

## Sideloading into Outlook (all methods)

Once your manifest URL is live, install the add-in:

### Outlook on the web (outlook.office.com)

1. Open any email → click **...** (more options) in the ribbon
2. Select **Get Add-ins**
3. Go to **My add-ins** → **Add a custom add-in** → **Add from URL**
4. Paste: `https://YOUR-DOMAIN/manifest.xml`
5. Click **Add** → confirm the permissions prompt

### Outlook desktop (Windows / Mac)

1. Open Outlook → **File** → **Manage Add-ins** (opens browser)
2. Same steps as above — go to My add-ins → Add from URL

### Microsoft 365 Admin Center (deploy to whole org)

> Requires M365 admin access

1. Sign in to [admin.microsoft.com](https://admin.microsoft.com)
2. **Settings** → **Integrated apps** → **Upload custom apps**
3. Choose **Provide link to manifest** → paste your manifest URL
4. Assign to users/groups → **Deploy**

---

## Permissions explained

The manifest requests `ReadWriteMailbox`. Here's what that enables:

| Permission | What it's used for |
|---|---|
| Read inbox messages | Scanning for unanswered emails |
| Read sent items | Finding threads with no reply |
| Send mail as user | Sending replies from the nudge panel |
| Roaming settings | Saving your nudge threshold + toggles |

The REST token is short-lived (1 hour) and obtained fresh each scan. No credentials are stored anywhere in the app.

---

## Customising the nudge threshold

Default: flag emails with no reply after **2 days**.

Change via the settings gear in the add-in, or update the default in `taskpane.html`:

```js
// In getSettings():
nudgeDays: s.get('nudgeDays') || 2,   // ← change this number
```

---

## Local development (with ngrok)

Office.js requires HTTPS even for local testing. Use ngrok to tunnel localhost:

```bash
npm install
node server.js &          # starts on port 3000

# In a second terminal:
npx ngrok http 3000       # gives you https://abc123.ngrok.io
```

Replace `YOUR-DOMAIN` in `manifest.xml` with the ngrok URL, then sideload as above. The ngrok URL changes each restart (paid plan fixes it).

---

## Troubleshooting

**"Add-in could not be loaded"**
→ Check that your server is returning HTTPS (not HTTP). Office refuses to load HTTP content.

**Nudge list is empty after scanning**
→ The REST token request may have failed. Open browser DevTools (F12 in Outlook web), check the Console for fetch errors. Common cause: the add-in domain isn't listed in `<AppDomains>` in the manifest.

**"Reply failed"**
→ Confirm `ReadWriteMailbox` permission is in the manifest (it is by default). Some orgs restrict third-party send permissions — check with your M365 admin.

**Settings not persisting**
→ `roamingSettings` syncs to Exchange and can take 30–60 seconds. If it never persists, check that the mailbox version supports Mailbox 1.5 (required).

**Add-in not showing in ribbon**
→ Close and reopen Outlook. New add-ins sometimes require a full restart to appear in the ribbon.
