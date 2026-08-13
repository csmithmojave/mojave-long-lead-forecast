# Mojave HVAC — Demand Forecast Dashboard

## Deploy to Railway

1. Push this folder to a GitHub repository
2. In Railway: **New Project → Deploy from GitHub repo** → select the repo
3. Railway auto-detects the `Procfile` and deploys
4. Open the generated Railway URL — done

## Files

| File | Purpose |
|------|---------|
| `index.html` | The entire dashboard (self-contained) |
| `Procfile` | Tells Railway to run `npx serve` |
| `package.json` | Node metadata + `serve` dependency |
| `.gitignore` | Excludes node_modules |

## Updating

Replace `index.html` with the latest version and push to GitHub.  
Railway auto-redeploys on every push.

## localStorage

On Railway (HTTPS), the dashboard saves uploaded CRM + planning data between sessions automatically.  
Data is stored under these keys:
- `mojave-dash-data` — CRM forecast + scheduled demand
- `mojave-open-lines` — open order quantities
- `mojave-manual-demand` — manually entered demand
- `mojave-range` — date range selection
