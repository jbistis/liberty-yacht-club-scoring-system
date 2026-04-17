# Liberty Yacht Club Racing Scoring System

Automated PHRF scoring system for the Liberty Yacht Club Wednesday Night Race Series. This repository contains the Google Apps Script scoring engine and the Wix website display layer for each season.

---

## Repository Structure

```
liberty-yacht-club-scoring-system/
├── 2025/
│   ├── LYC-Racing-Scoring-2025.gs   # Google Apps Script - 2025 scoring rules
│   └── sheet-table.html             # Wix iframe HTML - season display
├── 2026/
│   ├── LYC-Racing-Scoring-2026.gs   # Google Apps Script - 2026 scoring rules
│   └── sheet-table.html             # Wix iframe HTML - season display (with year picker)
└── README.md
```

---

## How It Works

### Google Apps Script (`*.gs`)
Each season has its own Google Sheet (owned by info@libertyyachtclub.org) with a bound Apps Script project. The script:
- Reads finish times from the **Race Results Entry** sheet
- Calculates elapsed and corrected times using PHRF TCF formula
- Assigns points and places per race and class
- Generates **Calculated Results**, **Series Standings**, **Season Standings**, and **Cumulative Results** sheets automatically
- Exposes a `doGet` web app endpoint that serves sheet data as CSV to the Wix website

To recalculate results, open the Google Sheet and use the **⛵ Racing Scoring** menu → **Calculate All Results**.

---

## End-of-Race Workflow

At the end of each race, do the following in the Google Sheet:

1. Open the **Race Results Entry** tab.
2. Fill in one row per boat with race data in **columns A through M, and column O**:
   - **A** Race#
   - **B** Series
   - **C** RaceType
   - **D** Class
   - **E** Course
   - **F–G** Start Date / Start Time
   - **H** StartDateTime *(auto-computed from F + G)*
   - **I** Wind
   - **J** Tide
   - **K** BoatName
   - **L–M** Finish Date / Finish Time
   - **N** FinishDateTime *(auto-computed from L + M — leave blank)*
   - **O** Status *(leave blank for a clean finish, otherwise enter a code: OCS, DNF, RET, DNC, DSQ, DNE, TLE, BYE, etc.)*
3. Run the **⛵ Racing Scoring** menu → **Calculate All Results**.
4. Verify the **Calculated Results**, **Series Standings**, **Season Standings**, and **Cumulative Results** tabs updated correctly.

The Wix display page will reflect the new data on the next page load (data is fetched live from the Apps Script proxy).

### HTML Display (`sheet-table.html`)
A self-contained HTML/CSS/JavaScript file embedded in a Wix HTML iframe on the Racing Results page. It:
- Fetches data from the deployed Apps Script proxy URL
- Displays tabbed views of each results sheet
- Includes a **season year picker** (2025, 2026, etc.) that switches between seasons
- Supports search, column filtering, and sortable columns
- Is responsive for mobile and desktop

---

## Installing sheet-table.html in Wix

The HTML file is embedded as a **Wix HTML iframe** element on the Racing page (`libertyyachtclub.org/racing`).

### First-time installation
1. Open the **Wix Editor** for the Racing page
2. Click **Add Elements (+)** on the left sidebar
3. Go to **Embed & Social → HTML iframe**
4. Drag it onto the page below the Race Records section
5. Click the iframe element → **"Enter Code"**
6. Paste the entire contents of `sheet-table.html` into the code box
7. Click **"Update"**
8. Resize the iframe to full width and approximately **900–1000px tall**
9. Click **"Publish"** to push changes live

### Updating the HTML (e.g. new season URL)
1. Open the **Wix Editor** for the Racing page
2. Click on the existing HTML iframe element
3. Click **"Enter Code"**
4. Replace the existing code with the updated `sheet-table.html` contents
5. Click **"Update"** then **"Publish"**

### Updating the SEASONS config only
If you only need to add a new season URL, find this block near the top of the script section in `sheet-table.html` and update it:

```javascript
const SEASONS = {
  2026: 'https://script.google.com/macros/s/YOUR_2026_URL/exec',
  2025: 'https://script.google.com/macros/s/YOUR_2025_URL/exec',
};
const DEFAULT_YEAR = 2026; // ← update this each new season
```

### Important notes
- The Wix HTML iframe has a size limit — keep the HTML file lean and self-contained
- The Apps Script deployment must be set to **"Anyone"** access or the iframe will fail to fetch data
- The **"Anyone with the link - Viewer"** setting on the Google Sheet must remain enabled
- Do **not** use the Wix mobile editor to resize the iframe — always use the desktop editor

---

## Season Setup Checklist (each new year)

1. **Copy** the previous year's Google Sheet into the LYC Google Drive
2. **Rename** the sheet to `Liberty YC Racing YYYY - Automated System`
3. **Clear** race data from Race Results Entry (keep Scratch Sheet boat registry)
4. **Open** Extensions → Apps Script and update the script with any new scoring rules
5. **Deploy** the script as a Web App (Execute as: Me, Who has access: Anyone)
6. **Copy** the deployed URL
7. **Update** `sheet-table.html` — add the new year's URL to the `SEASONS` object and update `DEFAULT_YEAR`
8. **Paste** the updated HTML into the Wix iframe on the Results page
9. **Update** the Wix automation trigger to point to the new year's registration form

---

## Scoring Rules

### PHRF Time Correction
```
TCF = 650 / (550 + PHRF Rating)
Corrected Time = Elapsed Time × TCF
```

### Points
| Finish | Points |
|--------|--------|
| 1st place | 1 |
| 2nd place | 2 |
| nth place | n |
| DNF / RET / OCS | Finishers + 1 |
| DNC / DSQ / DNE | Starters + 2 |
| TLE | Finishers + 2 |
| BYE (2025) | 0 |
| BYE (2026+) | Average of boat's other races in that series |

### Throwouts
One throwout for every 7 races sailed.

### Season Qualification
Boats must participate in at least **75%** of race days in the season to qualify for season standings.

---

## 2026 Rule Changes
- **BYE scoring**: BYE races are now scored as the average of the boat's other races within the same series (including DNC, RET, etc. in the average). BYE still counts toward the qualification threshold.

---

## Key URLs

| Item | Details |
|------|---------|
| 2025 Apps Script | Deployed from Liberty YC Racing 2025 - Automated System |
| 2026 Apps Script | Deployed from Liberty YC Racing 2026 - Automated System |
| Wix Results Page | libertyyachtclub.org/racing (Results section) |
| Wix Registration Page | libertyyachtclub.org/racing-registration |

---

## Maintainer Notes

- The `sheet-table.html` file is **identical across seasons** except for the `SEASONS` config block at the top of the script section. Only that block needs updating each year.
- Each year's `.gs` file should be kept **independent** — do not share a single script across seasons. This preserves the ability to rescore historical seasons under the rules that applied at the time.
- The Apps Script `doGet` function uses `SpreadsheetApp.getActiveSpreadsheet()` — no hardcoded sheet IDs. The script always serves data from whichever sheet it is bound to.
- To edit and deploy the Apps Script, you must be added as an Editor on the script project via **Share Sheet + Script** from script.google.com by the LYC account owner.
