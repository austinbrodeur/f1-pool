# AGENTS.md

This file provides guidance to AI agents when working with code in this repository.

## Project Overview

This is a Formula 1 prediction pool system built on Google Apps Script + Google Sheets. Players predict pole position and top 10 finishers for each race. The system uses a custom web app for pick submission instead of Google Forms to enable real-time duplicate prevention and better UX.

**Critical:** This is a **sheet-bound Apps Script project**. It MUST be created via Extensions → Apps Script from within a Google Sheet. Standalone projects won't work because the code relies on `SpreadsheetApp.getActiveSpreadsheet()` and custom menus only appear in bound projects.

## Architecture

### Three-File Structure

1. **SheetsScript.gs** - Main backend
   - Setup, lock/import, scoring, results emails, leaderboard
   - Custom menu system accessed via "F1 Pool" menu in Google Sheets
   - Configuration constants for drivers, teams, races, scoring rules

2. **WebScript.gs** (also named WebApp.gs) - Web app layer
   - Serves the pick page via `doGet()`
   - Token-based authentication (base64)
   - Server-side functions called from HTML: `getPickPageData()`, `submitPicks()`
   - **Must update `WEB_APP_URL` constant after deployment**

3. **PickPage.html** - Player-facing UI
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
Admin scores race (enters top 15 + pole)
  ↓
System emails personalized results to all players
```

## Deployment Workflow

### Initial Setup
1. Create Google Sheet
2. Extensions → Apps Script
3. Add three files: SheetsScript.gs, WebScript.gs, PickPage.html
4. Run `initialSetup` function (creates sheets, menu appears)
5. Deploy → New deployment → Web app → "Anyone can access"
6. Copy deployed URL and update `WEB_APP_URL` in WebScript.gs
7. **Redeploy** after updating the URL

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
- Calculated as `raceDate - 3 days` (Thu 8PM for Sun races, Wed 8PM for Sat races)
- Enforced client-side (countdown, disabled submit) AND server-side (rejects late submissions)

### Email Styling
Results and pick notification emails use dark theme matching the web app:
- Full HTML structure with `<body style="background-color:#0d0d0d !important">`
- This forces dark background in email clients (Gmail strips standalone `<body>` tags)

## Admin Workflow Per Race

Access via "F1 Pool" menu in Google Sheets:

1. **Picks → Create & Email Web Picks** - Generates unique URLs, emails all players
2. Wait for deadline (auto-enforced)
3. **Lock & Import → Lock Picks** - Imports responses, validates, applies carry-overs
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

All configuration is in SheetsScript.gs constants:
- `DRIVERS_2026` - Array of 3-letter driver codes
- `DRIVER_NAMES` - Map of code → last name
- `TEAMS_2026` - Map of code → team name
- `RACES_2026` - Array of race objects (round, name, date, day)
- `SCORING` - Points array for position accuracy

## Known Gotchas

1. **HTML template escaping:** `<?= ?>` escapes HTML entities. For JSON injection, ALWAYS use `<?!= JSON.stringify(value) ?>` (raw output).

2. **Can't make standalone project sheet-bound:** If you accidentally create a standalone project, you must start over from the Google Sheet.

3. **Gmail limits:** ~100 emails/day (regular accounts), 1500/day (Workspace). Current pool has 42 players.

4. **Token secret persistence:** Stored in Script Properties. If deleted/changed, all existing player links break. Never regenerate unless intentional.

5. **Race calendar gaps:** System handles non-sequential round numbers. Safe to remove cancelled races from `RACES_2026` array without reindexing.

## Color Scheme

Web app and emails use dark F1 theme:
- Background: `#0d0d0d`
- Card: `#161616`
- Accent green: `#00d4aa` (teal)
- Red: `#e10600` (F1 red)
- Gold: `#d4a017` (deadline/rankings)
- Pole purple: `#9b59b6` (F1 fastest lap)
- Team colors: Defined in CSS variables via `.t-{teamname}` classes
