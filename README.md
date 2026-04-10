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

### HTML Display (`sheet-table.html`)
A self-contained HTML/CSS/JavaScript file embedded in a Wix HTML iframe on the Racing Results page. It:
- Fetches data from the deployed Apps Script proxy URL
- Displays tabbed views of each results sheet
- Includes a **season year picker** (2025, 2026, etc.) that switches between seasons
- Supports search, column filtering, and sortable columns
- Is responsive for mobile and desktop

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
