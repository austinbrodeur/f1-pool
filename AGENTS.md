# AGENTS.md

This file provides guidance to AI agents when working with code in this repository.

## Project Overview

This is a Formula 1 prediction pool system built on Google Apps Script + Google Sheets. Players predict pole position and top 10 finishers for each race. The system uses a custom web app for pick submission instead of Google Forms to enable real-time duplicate prevention and better UX.

**Critical:** This is a **sheet-bound Apps Script project**. It MUST be created via Extensions → Apps Script from within a Google Sheet. Standalone projects won't work because the code relies on `SpreadsheetApp.getActiveSpreadsheet()` and custom menus only appear in bound projects.

## Architecture

### Five-File Structure

1. **SeasonConfig.gs** - Pre-initialization bootstrap configuration
   - Contains admin-editable JSON for the season year, pool name, timezone, deadline rules, drivers, teams, and races
   - Used only by `initialSetup()`; after setup, the generated sheets are the live source of truth

2. **SheetsScript.gs** - Main backend
   - Setup, lock/import, scoring, results emails, leaderboard
   - Custom menu system accessed via "F1 Pool" menu in Google Sheets
   - Runtime configuration helpers read Settings, Teams, Drivers, and Race Calendar

3. **PublicResults.gs** - Viewer-safe results workbook
   - Creates and remembers a separate Google spreadsheet via Script Properties
   - Publishes locked-pick tabs, race score tabs, Summary, Our Race, and Points Data Graph
   - Refreshes after leaderboard recalculation and before results emails
   - Adds player emails as viewers and attempts anyone-with-link viewing

4. **WebScript.gs** (also named WebApp.gs) - Web app layer
   - Serves the pick page via `doGet()`
   - Token-based authentication (base64)
   - Server-side functions called from HTML: `getPickPageData()`, `submitPicks()`
   - **Must update `WEB_APP_URL` constant after deployment**

5. **PickPage.html** - Player-facing UI
   - Dark F1-themed single-page app
   - Desktop drag-and-drop + mobile touch handling
   - Client-side duplicate prevention and swap logic
   - Visual feedback: P1=gold border, P2=silver, P3=bronze, pole=purple text

### Data Flow

```
Player receives unique URL with token
  ↓
WebScript.doGet() validates token, serves PickPage.html
  ↓
PickPage calls getPickPageData() via google.script.run
  ↓
Player drags/drops drivers, submits picks
  ↓
submitPicks() validates and writes to hidden response sheet
  ↓
Admin locks picks via menu (imports from response sheet)
  ↓
PublicResults publishes the locked picks and emails all players the viewer link
  ↓
Admin scores race (enters top 15 + pole)
  ↓
PublicResults refreshes the viewer-safe workbook
  ↓
System emails personalized results to all players
```

## Deployment Workflow

### Initial Setup
1. Create Google Sheet
2. Extensions → Apps Script
3. Add five files: SeasonConfig.gs, SheetsScript.gs, PublicResults.gs, WebScript.gs, PickPage.html
4. Edit the JSON in SeasonConfig.gs for the upcoming season
5. Run `initialSetup` function (creates sheets, menu appears)
6. Deploy → New deployment → Web app → "Anyone can access"
7. Copy deployed URL and update `WEB_APP_URL` in WebScript.gs
8. **Redeploy** after updating the URL

### Making Changes
After editing any .gs or .html file:
- Deploy → Manage deployments → Edit → **New version**
- Just saving the script does NOT update the live web app

## Key Technical Details

### Token System
- Format: `playerName|roundNum|secret` (base64-encoded)
- Secret is UUID stored in Script Properties (never changes)
- Tokens are player-specific and round-specific
- **Do NOT use `<?= ?>` in HTML templates** - it HTML-escapes output and breaks JSON. Use `<?!= ?>` for raw output

### Duplicate Prevention
The system prevents duplicates in P1-P10 (pole is separate and can duplicate):
- Client-side: Visual indicators + swap validation in `dropOn()`
- Swaps within P1-P10: Always allowed
- Swaps involving pole: Validated to prevent duplicates in P1-P10
- Server-side: Final validation in `submitPicks()`

### Visual Feedback (PickPage.html)
Driver chips in the pool show position via CSS classes:
- `.p1` - Gold border (#FFD700)
- `.p2` - Silver border (#C0C0C0)
- `.p3` - Bronze border (#CD7F32)
- `.in-top10` - Team color border (P4-P10)
- `.pole` - Purple text (#9b59b6, F1's fastest lap color)

### Deadlines
- Calculated from the configurable deadline offset, hour, and timezone in Settings
- Enforced client-side (countdown, disabled submit) AND server-side (rejects late submissions)

### Email Styling
Results and pick notification emails use dark theme matching the web app:
- Full HTML structure with `<body style="background-color:#0d0d0d !important">`
- This forces dark background in email clients (Gmail strips standalone `<body>` tags)

## Admin Workflow Per Race

Access via "F1 Pool" menu in Google Sheets:

1. **Picks → Create & Email Web Picks** - Generates unique URLs, emails all players
2. Wait for deadline (auto-enforced)
3. **Lock & Import → Lock Picks** - Imports responses, validates, applies carry-overs, publishes picks, and emails the public link
4. **Scoring → Score a Race** - Enter top 15 driver codes + pole + partial race multiplier
5. System prompts to email results (can resend later via menu)
6. Leaderboard auto-updates

## Scoring Rules

- **Pole:** 5 points for correct prediction
- **Position points** (by places off): Exact=25, ±1=18, ±2=15, ±3=12, ±4=10, ±5=8, ±6=6, ±7=4, ±8=2, ±9=1
- No points for drivers finishing outside top 15
- Partial races: Points × (laps completed / laps scheduled)
- **Carry-over:** If no valid picks submitted, uses previous race picks (marked with *)

## Configuration

- Before initialization, edit the JSON payload in `SeasonConfig.gs`.
- `initialSetup()` copies it into the Settings, Teams, Drivers, and Race Calendar sheets.
- After initialization, those four sheets are the runtime source of truth; edit them for live season changes.
- Scoring constants remain in `SheetsScript.gs`.

## Known Gotchas

1. **HTML template escaping:** `<?= ?>` escapes HTML entities. For JSON injection, ALWAYS use `<?!= JSON.stringify(value) ?>` (raw output).

2. **Can't make standalone project sheet-bound:** If you accidentally create a standalone project, you must start over from the Google Sheet.

3. **Gmail limits:** ~100 emails/day (regular accounts), 1500/day (Workspace). Current pool has 42 players.

4. **Token secret persistence:** Stored in Script Properties. If deleted/changed, all existing player links break. Never regenerate unless intentional.

5. **Race calendar gaps:** System handles non-sequential round numbers. Races can be removed or reordered in the Race Calendar sheet without reindexing.

6. **Bootstrap vs. live config:** Editing `SeasonConfig.gs` after initialization does not update existing sheets. Change the Settings, Teams, Drivers, or Race Calendar sheets instead.

7. **Public results workbook:** Its spreadsheet ID is stored in the `public_results_spreadsheet_id` Script Property. Deleting that property causes the next refresh to create a new workbook; old emailed links will then continue pointing to the previous file.

8. **Existing pools without Teams:** Use F1 Pool → Utilities → Create Teams Config Sheet. This adds the live color configuration without rerunning destructive initialization.

9. **Locked-picks email retries:** Successful recipients are stored per round in the `locked_picks_emailed_R{round}` Script Property. Re-locking retries failed recipients without emailing successful recipients again.

10. **Public link access:** Generated links include the Drive resource key when one is required. Use F1 Pool → Utilities → Check Public Results Access to verify link-sharing status and get an incognito-test URL. Workspace policy can still prohibit anyone-with-link sharing; named player viewers are added as a fallback.

## Color Scheme

Web app and emails use dark F1 theme:
- Background: `#0d0d0d`
- Card: `#161616`
- Accent green: `#00d4aa` (teal)
- Red: `#e10600` (F1 red)
- Gold: `#d4a017` (deadline/rankings)
- Pole purple: `#9b59b6` (F1 fastest lap)
- Team colors: Seeded from `SeasonConfig.gs`, stored in the Teams sheet, and sent to the web UI with driver data
