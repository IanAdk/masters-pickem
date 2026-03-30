# Masters Pick'em 2026 — Setup Guide

## Overview
- **`index.html`** — hosted on GitHub Pages (static, zero cost)
- **`Code.gs`** — Google Apps Script web app (backend: reads/writes picks to a Google Sheet)

---

## Step 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) → New blank spreadsheet
2. Name it: `Masters Pickems 2026`
3. Copy the Sheet ID from the URL:
   `https://docs.google.com/spreadsheets/d/` **`<THIS_PART>`** `/edit`

---

## Step 2 — Deploy the Apps Script

1. In your Google Sheet: **Extensions → Apps Script**
2. Delete any existing code, paste in the contents of `Code.gs`
3. Set `SHEET_ID` at the top to your Sheet ID from Step 1
4. Click **Save** (💾)
5. Run `initSheet()` once to create the tabs and seed all 101 golfers:
   - Click the function dropdown (top of editor) → select `initSheet`
   - Click ▶ Run — approve permissions when prompted
6. **Deploy as web app:**
   - Click **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy** → copy the web app URL

---

## Step 3 — Configure the frontend

Open `index.html` and replace the placeholder near the top of the `<script>` block:

```js
const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
```

Paste in the URL from Step 2.

---

## Step 4 — Host on GitHub Pages

1. Create a new GitHub repo (e.g. `masters-pickems`)
2. Push `index.html` to the `main` branch
3. Go to repo **Settings → Pages → Branch: main / root → Save**
4. Your URL: `https://<your-username>.github.io/masters-pickems/`

Share that link with all players!

---

## Step 5 — ESPN Event ID (for live scores)

The ESPN event ID for The Masters 2026 won't be known until close to the tournament.

**To find it (around April 7–9):**
1. Visit: `https://site.api.espn.com/apis/site/v2/sports/golf/leaderboard?league=pga`
2. Look in the JSON for the Masters event — copy its `id` value
3. On the live site, click the **Set** button next to "ESPN Event ID" and paste it in

The ID is saved in browser localStorage so each viewer only needs to set it once (or you can hardcode it in `index.html` before the tournament).

---

## Scoring Rules

| Situation | Score counted |
|-----------|---------------|
| Main golfer makes cut | Their score (relative to par) |
| Main golfer misses cut | Alternate's score used instead (permanently) |
| Multiple mains miss cut | Still only one alternate |

- **Lower is better** (standard golf scoring vs. par)
- **Tiebreaker:** closest birdie guess to actual total tournament birdies
- **Pick lock:** April 9, 2026 at 6:00 AM ET (hardcoded; auto-enforced in browser)

---

## Admin: Set Actual Birdies (after tournament)

To resolve tiebreakers, POST to your Apps Script:

```js
fetch(APPS_SCRIPT_URL, {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ action: 'updateBirdies', birdies: 284, adminKey: 'masters2026' })
});
```

Change the admin key: in Apps Script UI → Project Settings → Script Properties → add `ADMIN_KEY`.
