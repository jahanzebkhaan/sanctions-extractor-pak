# Global Sanctions Lists — Railway Deployment

Hosts a web app that downloads all 5 sanctions lists daily and serves
a consolidated Excel file via a single download link.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask web app — dashboard + download endpoint |
| `sanctions_extractor.py` | The extractor (all 5 parsers) |
| `requirements.txt` | Python dependencies |
| `Dockerfile` | Container with Playwright/Chromium pre-installed |

## Deploy to Railway (10 minutes)

### Step 1 — GitHub repo
1. Create a new repo on github.com (e.g. `sanctions-extractor`)
2. Upload all 4 files into it

### Step 2 — Railway account
1. Go to [railway.app](https://railway.app)
2. Sign up free (GitHub login recommended)

### Step 3 — Deploy
1. Click **New Project** → **Deploy from GitHub repo**
2. Select your `sanctions-extractor` repo
3. Railway auto-detects the Dockerfile and builds it
4. Wait ~3 minutes for build to complete

### Step 4 — Get your URL
1. Click your service → **Settings** → **Networking**
2. Click **Generate Domain**
3. Your URL: `https://sanctions-extractor-xxxx.railway.app`

## Usage

| URL | Action |
|-----|--------|
| `https://your-app.railway.app/` | Dashboard — status + record counts |
| `https://your-app.railway.app/download` | Download latest Excel |
| `https://your-app.railway.app/refresh` | Trigger manual refresh |
| `https://your-app.railway.app/status` | JSON status |

## Auto-refresh
Lists refresh automatically every day at midnight UTC.
First run happens automatically when the app starts.

## NACTA Note
Full 4,779 NACTA records require the Playwright Chromium download.
The Dockerfile pre-installs Chromium so this works on Railway.

## Cost
Railway free tier: $5 credit/month.
This app uses ~$1-2/month → effectively free.
