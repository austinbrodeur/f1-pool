# Pre-Season TODOs

## Season configuration

- [x] Move the season year, pool details, drivers, teams, and race calendar into the bootstrap JSON in `SeasonConfig.gs`.
  - `initialSetup()` populates Settings, Drivers, and Race Calendar from the bootstrap file.
  - The generated sheets are the live source of truth after initialization.
  - Calendar status updates locate rows by round number rather than assuming `roundNum + 1`.

## Public results

- [x] Recreate the legacy Summary, Our Race, and Points Data Graph views in a separate viewer-safe workbook.
  - Include public copies of every locked-picks and scored-race sheet so players can review all picks.
  - Email the public picks link to all players immediately after locking.
  - Refresh automatically after scoring and include public race/master links in results emails.
  - Grant listed player emails viewer access and request anyone-with-link viewing when permitted by Google Workspace policy.

> **Setup warning:** `initialSetup()` clears and rebuilds the Settings, Teams, Drivers, Race Calendar, Players, and Leaderboard sheets. Do not run it on an active pool without intentionally resetting that data.
