# F1 Pool 2026 — Project Context

## Overview

A prediction pool for the 2026 Formula 1 season with 42 players. Each race, players predict the top 10 finishers and pole position. Points are awarded based on accuracy. Built on Google Sheets + Apps Script with a custom web app for pick submission.

## Architecture

Three files in a single **sheet-bound** Apps Script project (must be created via Extensions → Apps Script from a Google Sheet — standalone projects won't work since `SpreadsheetApp.getActiveSpreadsheet()` requires a bound project, and custom menus only appear in bound projects):

- **SheetsScript.gs** — Main backend: setup, lock/import, scoring, results emails, leaderboard
- **WebApp.gs** — Web app: serves the pick page, handles token auth, writes picks to sheet
- **PickPage.html** — Player-facing drag-and-drop pick interface

All Google Forms logic has been removed. The forms-based version exists as a backup but is not used.

## Current State

- **This system is for the 2027 season onward.** The 2026 season is being run on a legacy Excel-based system.
- A flexible migration tool (`F1Pool_migrate.gs`) exists but is not needed for a fresh start.
- The web app is deployed and functional.
- The pick page uses drag-and-drop with team grouping and colors.

## Key Technical Decisions

### Driver Codes
All drivers use 3-letter codes (VER, NOR, RUS, etc.) everywhere — forms, scoring input, pick sheets, leaderboard. The form dropdowns show "VER - Verstappen (Red Bull)" for readability.

### Race Calendar
Matches the official 2026 FIA calendar with 24 rounds. The existing pool spreadsheet column order was used as the reference. Key notes:
- R9 is Barcelona-Catalunya, R16 is Spanish GP (Madrid) — two separate Spanish races
- R13 is Hungarian GP (was missing in early versions)
- R17 Azerbaijan and R22 Las Vegas are Saturday races
- Bahrain (R4) and Saudi Arabia (R5) may be cancelled — if so, just remove from `RACES_2026` array without reindexing. The system handles gaps in round numbers.

### Pick Deadlines
- **Thursday 8PM** before Sunday races
- **Wednesday 8PM** before Saturday races
- Calculated as `raceDate - 3 days` for all races (this works because Sun-3=Thu, Sat-3=Wed)
- The web app enforces deadlines client-side (countdown timer, disabled submit) AND server-side (rejects submissions past deadline)

### Token System
Each player gets a unique URL to submit picks. Tokens use HMAC-SHA256 signatures (not base64 — base64 had persistent issues with Apps Script's decoder in the web app context):
- Format: `roundNum.hmacSignature.encodedPlayerName`
- Secret is a UUID stored in Script Properties (stable across all execution contexts)
- Tokens are verified server-side on every request

**Critical HTML template bug we discovered:** In Apps Script HTML templates, `<?= ?>` HTML-escapes output. When used with `JSON.stringify()`, the quotes become `&quot;` entities which end up as literal quote characters in JS strings. Must use `<?!= ?>` (raw output) for injecting JS values.

### Unique Player Links
- Each player gets a personal pre-filled URL with their name baked into the token
- Players cannot submit under someone else's name
- Links are stored in Script Properties for resending

### Scoring Rules (from pool rules PDF)
- **Pole bonus:** 5 points for correct pole prediction
- **Position scoring** (places off from prediction): Exact=25, 1=18, 2=15, 3=12, 4=10, 5=8, 6=6, 7=4, 8=2, 9=1
- **No points** for drivers finishing outside top 15
- **Partial races:** Points multiplied by (laps completed / laps scheduled), rounded to 1 decimal
- **Carry-over:** If a player doesn't submit or has invalid picks (duplicates), previous race picks are used (marked with *)
- **Duplicate prevention:** Web app makes it physically impossible to pick the same driver twice in P1-P10. Pole is separate.

### Rescoring
Running "Score a Race" on an already-scored round shows a warning and replaces existing scores. This handles post-race penalties or input errors.

### Results Email
After scoring, the admin is prompted to email results to all players. Each player gets a personalized email showing:
- Race winner ($25 prize)
- Their personal score and rank (highlighted)
- Top 10 scoreboard
- Actual race result
- Season standings (top 5 + their own position if outside top 5)
- Can also be resent via menu

### Leaderboard
Dynamic columns — a new column (R1, R2, etc.) is added for each scored race showing per-race points. Also shows total points, race wins, and best race score. Top 6 highlighted with gold/silver/bronze for the prize structure:
- 1st: $480, 2nd: $380, 3rd: $280, 4th: $180, 5th: $80, 6th: $50
- Per-race winner: $25 (ties split)

### Season Payout
- Entry fee: $50, pool admin gets free buy-in
- 2026 had 42 players

## Sheet Structure

| Sheet | Purpose |
|-------|---------|
| Drivers | 22-driver grid with 3-letter codes, names, teams |
| Race Calendar | All 24 races, dates, deadlines, pick/lock/score status |
| Players | Player names, emails, payment tracking |
| Leaderboard | Cumulative standings with per-race columns |
| _R{N} Responses (hidden) | Raw pick submissions per race (hidden) |
| R{N} Picks — {Race Name} | Locked picks visible after deadline |
| R{N} Scores | Scoring breakdown per player |

## Admin Workflow Per Race

1. **F1 Pool → Picks → Create & Email Web Picks** — generates unique links, emails all players
2. Wait for deadline (web app auto-enforces)
3. **F1 Pool → Lock & Import → Lock Picks** — imports responses, validates, applies carry-overs, closes submissions
4. **F1 Pool → Scoring → Score a Race** — enter top 15 codes + pole + partial multiplier
5. Prompted to email results to all players
6. Leaderboard auto-updates

## Web App Pick Page

- Dark F1-themed design with Barlow Condensed typography
- Drag-and-drop interface (desktop HTML5 drag + mobile touch handling)
- Drivers grouped by team with official team colors
- Used drivers fade out in the pool (impossible to pick duplicates)
- Live countdown timer to deadline
- Auto-disables at deadline (client + server enforcement)
- Loads previous submission for editing
- Progress bar showing picks made

## Player Roster (2026)

42 players — see the summary PDF for the full list with Round 1 scores. Notable: pool admin Austin Brodeur is one of the players.

## 2026 Driver Grid

| Code | Driver | Team |
|------|--------|------|
| VER | Verstappen | Red Bull |
| HAD | Hadjar | Red Bull |
| NOR | Norris | McLaren |
| PIA | Piastri | McLaren |
| RUS | Russell | Mercedes |
| ANT | Antonelli | Mercedes |
| LEC | Leclerc | Ferrari |
| HAM | Hamilton | Ferrari |
| SAI | Sainz | Williams |
| ALB | Albon | Williams |
| LAW | Lawson | Racing Bulls |
| LIN | Lindblad | Racing Bulls |
| ALO | Alonso | Aston Martin |
| STR | Stroll | Aston Martin |
| OCO | Ocon | Haas |
| BEA | Bearman | Haas |
| HUL | Hulkenberg | Audi |
| BOR | Bortoleto | Audi |
| GAS | Gasly | Alpine |
| COL | Colapinto | Alpine |
| PER | Perez | Cadillac |
| BOT | Bottas | Cadillac |

## Known Issues / Gotchas

1. **Apps Script HTML templates:** Use `<?!= ?>` not `<?= ?>` when injecting JS values via `JSON.stringify()`. The escaping version corrupts strings with literal quote characters.
2. **Apps Script base64:** `Utilities.base64Decode()` and `base64DecodeWebSafe()` have unexplained failures in the web app execution context. Token system was rebuilt to avoid base64 entirely.
3. **Standalone vs bound projects:** The script MUST be in a sheet-bound project. Can't retroactively bind a standalone project to a sheet.
4. **Web app redeployment:** After editing code, must create a new version (Deploy → Manage deployments → Edit → New version). Just saving the script doesn't update the live web app.
5. **`form.getShortenedUrl()`** was deprecated — removed from the script.
6. **Gmail daily limits:** ~100 emails/day for regular accounts, 1500 for Workspace. With 42 players this is fine.
7. **Token regeneration:** If the token secret in Script Properties is deleted or changes, all existing player links break. New links must be generated.

## Files Not In Use

- **F1Pool_migrate.gs** — Flexible migration tool (two-step: create template → import). Not needed for fresh season start.
- **Google Forms version** — Backed up separately. Was replaced by web app due to inability to prevent duplicate driver picks and desire for better UX.
- **F1_Pool_Setup_Guide.pdf** — Setup guide PDF. Will need updating to reflect web app changes.
