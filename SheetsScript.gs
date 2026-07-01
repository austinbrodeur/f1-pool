// =============================================================================
// 2026 FORMULA 1 POOL — Google Sheets Backend
// =============================================================================
// Pick submission is handled by the Web App (WebApp.gs + PickPage.html).
// This script manages: setup, lock/import, scoring, results emails, leaderboard.
//
// SETUP:
//   1. Create a Google Sheet
//   2. Extensions > Apps Script > paste this file + WebApp.gs + PickPage.html
//   3. Run "initialSetup" from the F1 Pool menu
//   4. Deploy the web app and update WEB_APP_URL in WebApp.gs
//   5. Use the "F1 Pool" menu for everything
// =============================================================================

// ── CONFIG ──────────────────────────────────────────────────────────────────

const DRIVERS_2026 = [
  "VER", "HAD", "NOR", "PIA",
  "RUS", "ANT", "LEC", "HAM",
  "SAI", "ALB", "LAW", "LIN",
  "ALO", "STR", "OCO", "BEA",
  "HUL", "BOR", "GAS", "COL",
  "PER", "BOT"
];

const DRIVER_NAMES = {
  "VER": "Verstappen",  "HAD": "Hadjar",     "NOR": "Norris",     "PIA": "Piastri",
  "RUS": "Russell",     "ANT": "Antonelli",  "LEC": "Leclerc",    "HAM": "Hamilton",
  "SAI": "Sainz",       "ALB": "Albon",      "LAW": "Lawson",     "LIN": "Lindblad",
  "ALO": "Alonso",      "STR": "Stroll",     "OCO": "Ocon",       "BEA": "Bearman",
  "HUL": "Hulkenberg",  "BOR": "Bortoleto",  "GAS": "Gasly",      "COL": "Colapinto",
  "PER": "Perez",       "BOT": "Bottas"
};

const TEAMS_2026 = {
  "VER": "Red Bull",    "HAD": "Red Bull",   "NOR": "McLaren",    "PIA": "McLaren",
  "RUS": "Mercedes",    "ANT": "Mercedes",   "LEC": "Ferrari",    "HAM": "Ferrari",
  "SAI": "Williams",    "ALB": "Williams",   "LAW": "Racing Bulls","LIN": "Racing Bulls",
  "ALO": "Aston Martin","STR": "Aston Martin","OCO": "Haas",      "BEA": "Haas",
  "HUL": "Audi",        "BOR": "Audi",       "GAS": "Alpine",     "COL": "Alpine",
  "PER": "Cadillac",    "BOT": "Cadillac"
};

const RACES_2026 = [
  { round: 1,  name: "Australian GP",          date: "2026-03-08", day: "Sun" },
  { round: 2,  name: "Chinese GP",             date: "2026-03-15", day: "Sun" },
  { round: 3,  name: "Japanese GP",            date: "2026-03-29", day: "Sun" },
  { round: 4,  name: "Bahrain GP",             date: "2026-04-12", day: "Sun" },
  { round: 5,  name: "Saudi Arabian GP",       date: "2026-04-19", day: "Sun" },
  { round: 6,  name: "Miami GP",               date: "2026-05-03", day: "Sun" },
  { round: 7,  name: "Canadian GP",            date: "2026-05-24", day: "Sun" },
  { round: 8,  name: "Monaco GP",              date: "2026-06-07", day: "Sun" },
  { round: 9,  name: "Barcelona-Catalunya GP",  date: "2026-06-21", day: "Sun" },
  { round: 10, name: "Austrian GP",            date: "2026-07-05", day: "Sun" },
  { round: 11, name: "British GP",             date: "2026-07-19", day: "Sun" },
  { round: 12, name: "Belgian GP",             date: "2026-07-19", day: "Sun" },
  { round: 13, name: "Hungarian GP",           date: "2026-07-26", day: "Sun" },
  { round: 14, name: "Dutch GP",               date: "2026-08-23", day: "Sun" },
  { round: 15, name: "Italian GP",             date: "2026-09-06", day: "Sun" },
  { round: 16, name: "Spanish GP",             date: "2026-09-13", day: "Sun" },
  { round: 17, name: "Azerbaijan GP",          date: "2026-09-26", day: "Sat" },
  { round: 18, name: "Singapore GP",           date: "2026-10-11", day: "Sun" },
  { round: 19, name: "United States GP",       date: "2026-10-25", day: "Sun" },
  { round: 20, name: "Mexico City GP",         date: "2026-11-01", day: "Sun" },
  { round: 21, name: "Brazilian GP",           date: "2026-11-08", day: "Sun" },
  { round: 22, name: "Las Vegas GP",           date: "2026-11-22", day: "Sat" },
  { round: 23, name: "Qatar GP",               date: "2026-11-30", day: "Sun" },
  { round: 24, name: "Abu Dhabi GP",           date: "2026-12-06", day: "Sun" }
];

const SCORING = [25, 18, 15, 12, 10, 8, 6, 4, 2, 1];
const POLE_BONUS = 5;
const TOP_FINISH_CUTOFF = 15;
const HEADER_BG = "#1a472a";
const HEADER_FG = "#ffffff";

// ── MENU ────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu("F1 Pool")
    .addItem("🏁 Initial Setup", "initialSetup")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("📧 Picks")
      .addItem("Create & Email Web Picks for Next Race", "createWebPicksForNextRace")
      .addItem("Create & Email Web Picks for Specific Race...", "createWebPicksPrompt")
      .addItem("Re-send Pick Link to All Players...", "resendWebLinksAll")
      .addItem("Re-send Pick Link to Specific Player...", "resendWebLinkToPlayer"))
    .addSubMenu(SpreadsheetApp.getUi().createMenu("🔒 Lock & Import")
      .addItem("Lock Picks for Next Unlocked Race", "lockPicksForNextRace")
      .addItem("Lock Picks for Specific Race...", "lockPicksPrompt"))
    .addSubMenu(SpreadsheetApp.getUi().createMenu("📊 Scoring")
      .addItem("Score a Race...", "scoreRacePrompt")
      .addItem("Resend Results Email for Race...", "emailRaceResultsPrompt")
      .addItem("Recalculate Leaderboard", "recalculateLeaderboard"))
    .addSubMenu(SpreadsheetApp.getUi().createMenu("🔧 Utilities")
      .addItem("View Deadline Status", "viewDeadlineStatus")
      .addItem("View Submission Status for Race...", "viewSubmissionStatus"))
    .addToUi();
}

// ── INITIAL SETUP ───────────────────────────────────────────────────────────

function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let ds = ss.getSheetByName("Drivers") || ss.insertSheet("Drivers");
  ds.clear();
  ds.getRange(1, 1, 1, 3).setValues([["Code", "Driver", "Team"]])
    .setFontWeight("bold").setBackground(HEADER_BG).setFontColor(HEADER_FG);
  DRIVERS_2026.forEach((code, i) => {
    ds.getRange(i + 2, 1).setValue(code);
    ds.getRange(i + 2, 2).setValue(DRIVER_NAMES[code]);
    ds.getRange(i + 2, 3).setValue(TEAMS_2026[code]);
  });
  ds.autoResizeColumns(1, 3);

  let cs = ss.getSheetByName("Race Calendar") || ss.insertSheet("Race Calendar");
  cs.clear();
  const calH = ["Round", "Race", "Race Date", "Day", "Picks Deadline",
                 "Picks Sent?", "Pick Link", "Locked?", "Scored?", "Race Winner (Pool)"];
  cs.getRange(1, 1, 1, calH.length).setValues([calH])
    .setFontWeight("bold").setBackground(HEADER_BG).setFontColor(HEADER_FG);
  RACES_2026.forEach((r, i) => {
    const row = i + 2;
    const deadline = getDeadline(r);
    cs.getRange(row, 1).setValue(r.round);
    cs.getRange(row, 2).setValue(r.name);
    cs.getRange(row, 3).setValue(r.date);
    cs.getRange(row, 4).setValue(r.day);
    cs.getRange(row, 5).setValue(deadline).setNumberFormat("MMM d, yyyy h:mm AM/PM");
    cs.getRange(row, 6).setValue("No");
    cs.getRange(row, 7).setValue("");
    cs.getRange(row, 8).setValue("No");
    cs.getRange(row, 9).setValue("No");
    cs.getRange(row, 10).setValue("");
  });
  cs.autoResizeColumns(1, calH.length);

  let ps = ss.getSheetByName("Players") || ss.insertSheet("Players");
  ps.clear();
  ps.getRange(1, 1, 1, 4).setValues([["Player Name", "Email", "Paid ($50)?", "Notes"]])
    .setFontWeight("bold").setBackground(HEADER_BG).setFontColor(HEADER_FG);
  ps.autoResizeColumns(1, 4);

  let lb = ss.getSheetByName("Leaderboard") || ss.insertSheet("Leaderboard");
  lb.clear();
  lb.getRange(1, 1, 1, 5)
    .setValues([["Rank", "Player", "Total Points", "Race Wins ($25)", "Best Race"]])
    .setFontWeight("bold").setBackground(HEADER_BG).setFontColor(HEADER_FG);

  const s1 = ss.getSheetByName("Sheet1");
  if (s1 && ss.getSheets().length > 1) try { ss.deleteSheet(s1); } catch(e) {}

  SpreadsheetApp.getUi().alert(
    "Setup Complete!",
    "Next steps:\n" +
    "1. Add player names & emails to the 'Players' sheet\n" +
    "2. Deploy the Web App (Deploy > New deployment > Web app)\n" +
    "3. Update WEB_APP_URL in WebApp.gs with the deployed URL\n" +
    "4. Use F1 Pool > Picks > Create & Email Web Picks for Next Race",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ── RESEND WEB PICK LINKS ───────────────────────────────────────────────────

function resendWebLinksAll() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Resend Links", "Enter round number:", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  sendWebLinksToPlayers(parseInt(r.getResponseText()));
}

function resendWebLinkToPlayer() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt("Resend Link", "Enter round number:", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const round = parseInt(r1.getResponseText());
  const r2 = ui.prompt("Player Email", "Enter the player's email:", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  sendWebLinksToPlayers(round, r2.getResponseText().trim());
}

function sendWebLinksToPlayers(roundNum, specificEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const race = RACES_2026.find(r => r.round === roundNum);
  if (!race) { ui.alert("Invalid round."); return; }

  const props = PropertiesService.getScriptProperties();
  const linksJson = props.getProperty(`form_links_${roundNum}`);
  if (!linksJson) { ui.alert(`No pick links found for Round ${roundNum}. Create them first.`); return; }

  const playerLinks = JSON.parse(linksJson);
  const deadline = getDeadline(race);
  const deadlineStr = Utilities.formatDate(deadline, Session.getScriptTimeZone(), "EEEE, MMM d 'at' h:mm a");
  const poolName = ss.getName();
  const players = getPlayers(ss);
  let sent = 0;

  players.forEach(player => {
    if (!player.email) return;
    if (specificEmail && player.email !== specificEmail) return;
    const playerUrl = playerLinks[player.name];
    if (!playerUrl) return;
    try {
      MailApp.sendEmail({
        to: player.email,
        subject: `🏎️ F1 Pool Picks — Round ${race.round}: ${race.name}` +
                 (specificEmail ? "" : " (Reminder)"),
        htmlBody: buildPickEmail(player.name, race, deadlineStr, playerUrl, poolName)
      });
      sent++;
    } catch(e) {}
  });

  ui.alert(`${sent} email(s) sent.`);
}

function buildPickEmail(playerName, race, deadlineStr, formUrl, poolName) {
  return `
    <html>
    <head>
      <meta name="color-scheme" content="light dark">
      <meta name="supported-color-schemes" content="light dark">
    </head>
    <body style="background-color: #0d0d0d !important; margin: 0; padding: 0;">
      <div style="background-color: #0d0d0d; margin: 0; padding: 20px; font-family: 'Barlow', -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;">
        <div style="max-width: 600px; margin: 0 auto; background: #161616; border-radius: 8px; overflow: hidden;">
          <div style="background: linear-gradient(135deg, #1a472a 0%, #0a1f12 100%); color: white; padding: 28px 20px; text-align: center; border-bottom: 3px solid #00d4aa;">
            <h1 style="margin: 0; font-size: 28px; font-weight: 800; text-transform: uppercase; letter-spacing: 3px;">🏎️ ${poolName}</h1>
            <p style="margin: 10px 0 0; font-size: 16px; opacity: 0.85; letter-spacing: 0.5px;">Round ${race.round}: ${race.name}</p>
          </div>
          <div style="background: #161616; padding: 32px 24px; color: #e8e8e8;">
            <p style="margin: 0 0 16px; font-size: 15px;">Hey <strong style="color: #00d4aa;">${playerName}</strong>,</p>
            <p style="margin: 0 0 20px; font-size: 15px; line-height: 1.6;">Time to submit your picks for the <strong>${race.name}</strong> on <strong>${race.date}</strong>.</p>
            <div style="background: rgba(212, 160, 23, 0.15); border-left: 4px solid #d4a017; border-radius: 6px; padding: 16px; margin: 20px 0;">
              <strong style="font-size: 15px; color: #d4a017;">⏰ Deadline: ${deadlineStr}</strong>
            </div>
            <p style="margin: 0 0 12px; font-size: 14px; line-height: 1.6; color: #b8b8b8;">Pick your <strong style="color: #e8e8e8;">Pole Position</strong> prediction and your predicted <strong style="color: #e8e8e8;">Top 10 finishing order</strong>.</p>
            <p style="margin: 0 0 24px; font-size: 14px; line-height: 1.6; color: #b8b8b8;">Each driver can only appear once in P1–P10. You can resubmit anytime before the deadline — only your latest picks count.</p>
            <div style="text-align: center; margin: 28px 0;">
              <a href="${formUrl}" style="background: #e10600; color: white; padding: 16px 40px; text-decoration: none; border-radius: 8px; font-size: 17px; font-weight: 700; display: inline-block; text-transform: uppercase; letter-spacing: 2px;">Submit Your Picks →</a>
            </div>
            <p style="color: #777; font-size: 12px; text-align: center; margin: 0;">If the button doesn't work, copy this link:<br><a href="${formUrl}" style="color: #00d4aa; word-break: break-all;">${formUrl}</a></p>
          </div>
          <div style="background: #0a0a0a; color: #777; padding: 16px; text-align: center; font-size: 12px;">
            2026 F1 Pool — Good luck! 🍀
          </div>
        </div>
      </div>
    </body>
    </html>
  `;
}

// ── LOCK & IMPORT PICKS ─────────────────────────────────────────────────────

function lockPicksForNextRace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cs = ss.getSheetByName("Race Calendar");
  if (!cs) return;
  const data = cs.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === "Yes" && data[i][7] === "No") {
      lockAndImportPicks(data[i][0]);
      return;
    }
  }
  SpreadsheetApp.getUi().alert("No races ready to lock.");
}

function lockPicksPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Lock Picks", "Enter round number:", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  lockAndImportPicks(parseInt(r.getResponseText()));
}

function lockAndImportPicks(roundNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const race = RACES_2026.find(r => r.round === roundNum);
  if (!race) { ui.alert("Invalid round."); return; }

  // Find the hidden responses sheet
  let respSheet = null;
  ss.getSheets().forEach(s => {
    if (s.getName().includes(`_R${roundNum} Responses`)) respSheet = s;
  });

  // Get all players
  const allPlayers = getPlayers(ss);
  if (allPlayers.length === 0) { ui.alert("No players found."); return; }

  // Parse responses — keep only latest per player
  const latestPicks = {};

  if (respSheet) {
    const respData = respSheet.getDataRange().getValues();
    // Column order: Timestamp, Player, Pole, P1..P10
    for (let i = 1; i < respData.length; i++) {
      const row = respData[i];
      const timestamp = row[0];
      const name = (row[1] || "").toString().trim();
      if (!name) continue;

      const pole = extractCode(row[2]);
      const picks = [];
      for (let c = 3; c <= 12; c++) {
        picks.push(extractCode(row[c]));
      }

      if (!latestPicks[name] || new Date(timestamp) > new Date(latestPicks[name].timestamp)) {
        latestPicks[name] = { pole, picks, timestamp };
      }
    }
  }

  // Create visible pick sheet
  const sheetName = `R${roundNum} Picks — ${race.name}`;
  let pickSheet = ss.getSheetByName(sheetName);
  if (pickSheet) pickSheet.clear();
  else pickSheet = ss.insertSheet(sheetName);

  const deadline = getDeadline(race);
  const deadlineStr = Utilities.formatDate(deadline, Session.getScriptTimeZone(), "EEEE, MMM d 'at' h:mm a");
  pickSheet.getRange(1, 1, 1, 14).merge()
    .setValue(`🔒 LOCKED — Round ${race.round}: ${race.name}  |  Deadline: ${deadlineStr}`)
    .setFontWeight("bold").setFontSize(12)
    .setBackground("#333333").setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  const headers = ["Player", "Status", "Pole", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "Submitted"];
  pickSheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground(HEADER_BG).setFontColor(HEADER_FG)
    .setHorizontalAlignment("center");

  let validCount = 0, carriedCount = 0, missingCount = 0, dupeCount = 0;
  let outputRow = 3;

  allPlayers.forEach((player, idx) => {
    const name = player.name;
    let pole = "", picks = Array(10).fill(""), status = "", submitted = "";
    let rowBg = idx % 2 === 0 ? "#ffffff" : "#f5f5f5";

    if (latestPicks[name]) {
      const data = latestPicks[name];
      pole = data.pole;
      picks = data.picks;
      submitted = Utilities.formatDate(new Date(data.timestamp), Session.getScriptTimeZone(), "MMM d, h:mm a");

      const filled = picks.filter(d => d);
      const unique = [...new Set(filled)];

      if (filled.length < 10) {
        const carried = getCarryOverPicks(ss, roundNum, name);
        if (carried) { pole = carried.pole; picks = carried.picks; status = "📋 Carried (incomplete)"; carriedCount++; }
        else { status = "⚠️ Incomplete"; missingCount++; }
      } else if (unique.length < filled.length) {
        dupeCount++;
        const carried = getCarryOverPicks(ss, roundNum, name);
        if (carried) { pole = carried.pole; picks = carried.picks; status = "📋 Carried (had dupes)"; carriedCount++; dupeCount--; }
        else { status = "⚠️ Had duplicates"; }
      } else {
        status = "✅ Valid"; validCount++;
      }
    } else {
      const carried = getCarryOverPicks(ss, roundNum, name);
      if (carried) { pole = carried.pole; picks = carried.picks; status = "📋 Carried over"; submitted = "—"; carriedCount++; }
      else { status = "❌ No picks"; submitted = "—"; missingCount++; }
    }

    pickSheet.getRange(outputRow, 1, 1, headers.length).setValues([[name, status, pole, ...picks, submitted]]);
    pickSheet.getRange(outputRow, 1, 1, headers.length).setBackground(rowBg);

    const statusCell = pickSheet.getRange(outputRow, 2);
    if (status.includes("✅")) statusCell.setFontColor("#007a33");
    else if (status.includes("📋")) statusCell.setFontColor("#cc8800");
    else statusCell.setFontColor("#cc0000");

    outputRow++;
  });

  pickSheet.setColumnWidth(1, 160); pickSheet.setColumnWidth(2, 150);
  for (let c = 3; c <= 13; c++) pickSheet.setColumnWidth(c, 140);
  pickSheet.setColumnWidth(14, 130);

  const protection = pickSheet.protect().setDescription("Locked picks");
  protection.setWarningOnly(true);

  // Update calendar
  const cs = ss.getSheetByName("Race Calendar");
  cs.getRange(roundNum + 1, 8).setValue("Yes");

  ss.setActiveSheet(pickSheet);

  ui.alert(
    `Picks Locked: ${race.name}`,
    `✅ Valid picks: ${validCount}\n📋 Carried over: ${carriedCount}\n` +
    `⚠️ Invalid: ${dupeCount}\n❌ No picks: ${missingCount}\n\n` +
    `Total players: ${allPlayers.length}\n\n` +
    `All picks are now visible in the "${sheetName}" sheet.`,
    ui.ButtonSet.OK
  );
}

function getCarryOverPicks(ss, currentRound, playerName) {
  if (currentRound <= 1) return null;
  for (let r = currentRound - 1; r >= 1; r--) {
    const sheetName = findSheet(ss, `R${r} Picks`);
    if (!sheetName) continue;
    const sheet = ss.getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      const name = (data[i][0] || "").toString().trim();
      if (name === playerName) {
        const pole = (data[i][2] || "").toString().trim();
        const picks = [];
        for (let c = 3; c <= 12; c++) picks.push((data[i][c] || "").toString().trim());
        const filled = picks.filter(d => d);
        const unique = [...new Set(filled)];
        if (filled.length === 10 && unique.length === 10) return { pole, picks };
      }
    }
  }
  return null;
}

// ── SCORING ─────────────────────────────────────────────────────────────────

function scoreRacePrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Score a Race", "Enter round number:", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const round = parseInt(r.getResponseText());
  const race = RACES_2026.find(rc => rc.round === round);
  if (!race) { ui.alert("Invalid round number, or this race has been removed from the calendar."); return; }
  scoreRace(round);
}

function scoreRace(roundNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const race = RACES_2026.find(r => r.round === roundNum);
  if (!race) { ui.alert("Race not found."); return; }

  // Rescore warning
  const existingScoreSheet = ss.getSheetByName(`R${roundNum} Scores`);
  if (existingScoreSheet) {
    if (ui.alert("⚠️ Rescore Warning",
      `Round ${roundNum} (${race.name}) has already been scored.\n` +
      "This will REPLACE the existing scores.\n\nContinue?",
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }

  const pickSheetName = findSheet(ss, `R${roundNum} Picks`);
  if (!pickSheetName) { ui.alert(`No locked pick sheet for Round ${roundNum}. Lock picks first.`); return; }
  const pickSheet = ss.getSheetByName(pickSheetName);

  // Results entry
  const resultPrompt = ui.prompt(`Results: ${race.name}`,
    "Enter the ACTUAL top 15 finishers using 3-letter codes, comma-separated (P1 through P15).\n\n" +
    "Example: RUS, ANT, LEC, HAM, NOR, VER, BEA, LIN, BOR, GAS, SAI, ALB, ALO, HUL, STR",
    ui.ButtonSet.OK_CANCEL);
  if (resultPrompt.getSelectedButton() !== ui.Button.OK) return;
  const actualResults = resultPrompt.getResponseText().split(",").map(s => s.trim().toUpperCase()).filter(s => s);
  if (actualResults.length < 10) { ui.alert("Enter at least 10 finishers."); return; }

  const unknown = actualResults.filter(d => !DRIVERS_2026.includes(d));
  if (unknown.length > 0) {
    if (ui.alert("Unknown Codes", `Not in grid: ${unknown.join(", ")}\nContinue?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }

  const poleResp = ui.prompt("Pole Position", "Who got pole? (3-letter code, e.g. VER)", ui.ButtonSet.OK_CANCEL);
  if (poleResp.getSelectedButton() !== ui.Button.OK) return;
  const actualPole = poleResp.getResponseText().trim().toUpperCase();

  const partialResp = ui.prompt("Partial Race?",
    "Full race = 1.0 | Partial = fraction (e.g. 0.73)\nBlank = full race.", ui.ButtonSet.OK_CANCEL);
  if (partialResp.getSelectedButton() !== ui.Button.OK) return;
  const multiplier = parseFloat(partialResp.getResponseText()) || 1.0;

  const resultsMap = {};
  actualResults.forEach((d, i) => { resultsMap[d] = i + 1; });

  const scoreSheetName = `R${roundNum} Scores`;
  let scoreSheet = ss.getSheetByName(scoreSheetName);
  if (!scoreSheet) scoreSheet = ss.insertSheet(scoreSheetName);
  scoreSheet.clear();

  const sHeaders = ["Player", "Pole Pick", "Pole Pts", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "Position Pts", "Raw Total", "Mult.", "TOTAL"];
  scoreSheet.getRange(1, 1, 1, sHeaders.length).setValues([sHeaders])
    .setFontWeight("bold").setBackground(HEADER_BG).setFontColor(HEADER_FG);

  const pickData = pickSheet.getDataRange().getValues();
  const scores = [];
  let outRow = 2;

  for (let i = 2; i < pickData.length; i++) {
    const playerName = (pickData[i][0] || "").toString().trim();
    if (!playerName) continue;
    const status = (pickData[i][1] || "").toString();
    if (status.includes("❌")) continue;

    const pole = (pickData[i][2] || "").toString().trim();
    const picks = [];
    for (let c = 3; c <= 12; c++) picks.push((pickData[i][c] || "").toString().trim());

    const isCarried = status.includes("📋");
    const polePts = (pole === actualPole) ? POLE_BONUS : 0;

    let posPts = 0;
    picks.forEach((driver, idx) => {
      const actual = resultsMap[driver];
      if (!actual || actual > TOP_FINISH_CUTOFF) return;
      const off = Math.abs(actual - (idx + 1));
      if (off < SCORING.length) posPts += SCORING[off];
    });

    const raw = polePts + posPts;
    const total = Math.round(raw * multiplier * 10) / 10;
    const displayName = isCarried ? playerName + " *" : playerName;

    scoreSheet.getRange(outRow, 1, 1, sHeaders.length).setValues([[
      displayName, pole, polePts, ...picks, posPts, raw, multiplier, total
    ]]);
    scores.push({ player: displayName, total });
    outRow++;
  }

  if (outRow > 2) scoreSheet.getRange(2, 1, outRow - 2, sHeaders.length).sort({ column: sHeaders.length, ascending: false });
  if (scores.length > 0) {
    scores.sort((a, b) => b.total - a.total);
    scoreSheet.getRange(2, 1, 1, sHeaders.length).setBackground("#d4edda");
  }

  // Actual results reference
  outRow = scores.length + 4;
  scoreSheet.getRange(outRow, 1).setValue("ACTUAL RESULTS").setFontWeight("bold");
  scoreSheet.getRange(outRow, 2).setValue("Pole: " + actualPole).setFontWeight("bold");
  actualResults.forEach((d, i) => {
    scoreSheet.getRange(outRow + 1 + i, 1).setValue(`P${i + 1}`);
    scoreSheet.getRange(outRow + 1 + i, 2).setValue(d);
  });

  scoreSheet.autoResizeColumns(1, sHeaders.length);

  const cs = ss.getSheetByName("Race Calendar");
  cs.getRange(roundNum + 1, 9).setValue("Yes");
  if (scores.length > 0) cs.getRange(roundNum + 1, 10).setValue(scores[0].player + " (" + scores[0].total + " pts)");

  recalculateLeaderboard();
  ss.setActiveSheet(scoreSheet);

  const carried = scores.filter(s => s.player.endsWith("*")).length;
  const sendEmail = ui.alert(
    `Scoring Complete: ${race.name}`,
    (scores.length > 0 ? `🏆 Winner: ${scores[0].player} — ${scores[0].total} pts\n\n` : "") +
    `${scores.length} players scored.` +
    (carried > 0 ? `\n${carried} used carried-over picks (*).` : "") +
    "\n\nLeaderboard updated.\n\nEmail results to all players?",
    ui.ButtonSet.YES_NO
  );

  if (sendEmail === ui.Button.YES) emailRaceResults(roundNum);
}

// ── EMAIL RACE RESULTS ──────────────────────────────────────────────────────

function emailRaceResultsPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Email Results", "Enter round number:", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  emailRaceResults(parseInt(r.getResponseText()));
}

function emailRaceResults(roundNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const race = RACES_2026.find(r => r.round === roundNum);
  if (!race) { ui.alert("Invalid round."); return; }

  const scoreSheet = ss.getSheetByName(`R${roundNum} Scores`);
  if (!scoreSheet) { ui.alert("No score sheet found."); return; }

  const players = getPlayers(ss);
  const playersWithEmail = players.filter(p => p.email);
  if (playersWithEmail.length === 0) { ui.alert("No players have email addresses."); return; }

  const scoreData = scoreSheet.getDataRange().getValues();
  const totalColIdx = scoreData[0].length - 1;

  const playerResults = [];
  let actualResultsStart = -1, actualPole = "";
  for (let i = 1; i < scoreData.length; i++) {
    const name = (scoreData[i][0] || "").toString().trim();
    if (!name) continue;
    if (name === "ACTUAL RESULTS") { actualResultsStart = i; actualPole = (scoreData[i][1] || "").toString().replace("Pole: ", ""); break; }
    playerResults.push({ name, total: parseFloat(scoreData[i][totalColIdx]) || 0, rank: playerResults.length + 1 });
  }

  const actualFinishers = [];
  if (actualResultsStart > 0) {
    for (let i = actualResultsStart + 1; i < scoreData.length; i++) {
      const pos = (scoreData[i][0] || "").toString().trim();
      const driver = (scoreData[i][1] || "").toString().trim();
      if (pos && driver) actualFinishers.push({ pos, driver });
    }
  }

  const lb = ss.getSheetByName("Leaderboard");
  const lbData = lb ? lb.getDataRange().getValues() : [];
  const standings = [];
  for (let i = 1; i < lbData.length; i++) {
    const name = (lbData[i][1] || "").toString().trim();
    if (name) standings.push({ rank: lbData[i][0], name, total: lbData[i][2] });
  }

  const poolName = ss.getName();
  const resultsSheetUrl = `${ss.getUrl()}#gid=${scoreSheet.getSheetId()}`;
  const raceWinner = playerResults.length > 0 ? playerResults[0] : null;
  let emailsSent = 0;

  playersWithEmail.forEach(player => {
    const myResult = playerResults.find(p => p.name.replace(/ \*$/, "") === player.name);
    const myStanding = standings.find(s => s.name === player.name);
    try {
      MailApp.sendEmail({
        to: player.email,
        subject: `🏁 F1 Pool Results — R${race.round}: ${race.name}`,
        htmlBody: buildResultsEmail(player.name, race, poolName, playerResults, myResult, myStanding, actualFinishers, actualPole, raceWinner, standings, resultsSheetUrl)
      });
      emailsSent++;
    } catch(e) { Logger.log(`Failed to email results to ${player.name}: ${e.message}`); }
  });

  ui.alert(`Results emailed to ${emailsSent} / ${playersWithEmail.length} players.`);
}

function buildResultsEmail(playerName, race, poolName, allResults, myResult, myStanding, actualFinishers, actualPole, raceWinner, standings, resultsSheetUrl) {
  const top10Rows = allResults.slice(0, 10).map((p, i) => {
    const isMe = p.name.replace(/ \*$/, "") === playerName;
    const bg = i === 0 ? "rgba(0, 212, 170, 0.15)" : (isMe ? "rgba(212, 160, 23, 0.15)" : (i % 2 === 0 ? "#1c1c1c" : "#161616"));
    const medal = i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : `${i + 1}.`;
    const borderColor = i === 0 ? "#00d4aa" : (isMe ? "#d4a017" : "#2a2a2a");
    return `<tr style="background:${bg};border-left:3px solid ${borderColor};"><td style="padding:8px 12px;border-bottom:1px solid #2a2a2a;color:#b8b8b8;">${medal}</td><td style="padding:8px 12px;border-bottom:1px solid #2a2a2a;${isMe ? "font-weight:bold;color:#00d4aa;" : "color:#e8e8e8;"}">${p.name}${isMe ? " ← You" : ""}</td><td style="padding:8px 12px;border-bottom:1px solid #2a2a2a;text-align:right;font-weight:bold;color:#e8e8e8;">${p.total}</td></tr>`;
  }).join("");

  const mySection = myResult
    ? `<div style="background:rgba(212, 160, 23, 0.15);border-left:4px solid #d4a017;border-radius:6px;padding:16px;margin:20px 0;"><strong style="color:#d4a017;font-size:16px;">Your Score: ${myResult.total} pts</strong> <span style="color:#b8b8b8;">— Ranked <strong style="color:#e8e8e8;">#${myResult.rank}</strong> of ${allResults.length}</span></div>`
    : `<div style="background:rgba(225, 6, 0, 0.15);border-left:4px solid #e10600;border-radius:6px;padding:16px;margin:20px 0;color:#e8e8e8;">You did not have scored picks for this race.</div>`;

  const resultsList = actualFinishers.slice(0, 10).map(f => `${f.pos}: <strong style="color:#e8e8e8;">${f.driver}</strong> <span style="color:#777;">${DRIVER_NAMES[f.driver] || f.driver}</span>`).join(" &nbsp;|&nbsp; ");

  const standingsRows = standings.slice(0, 5).map((s, i) => {
    const isMe = s.name === playerName;
    const bg = isMe ? "rgba(212, 160, 23, 0.15)" : (i % 2 === 0 ? "#1c1c1c" : "#161616");
    const borderColor = isMe ? "#d4a017" : "#2a2a2a";
    return `<tr style="background:${bg};border-left:3px solid ${borderColor};"><td style="padding:6px 10px;color:#b8b8b8;">${s.rank}</td><td style="padding:6px 10px;${isMe ? "font-weight:bold;color:#00d4aa;" : "color:#e8e8e8;"}">${s.name}${isMe ? " ←" : ""}</td><td style="padding:6px 10px;text-align:right;font-weight:bold;color:#e8e8e8;">${s.total}</td></tr>`;
  }).join("");

  let myStandingRow = "";
  if (myStanding && myStanding.rank > 5) {
    myStandingRow = `<tr style="background:rgba(212, 160, 23, 0.15);border-left:3px solid #d4a017;border-top:2px dashed #444;"><td style="padding:6px 10px;color:#b8b8b8;">${myStanding.rank}</td><td style="padding:6px 10px;font-weight:bold;color:#00d4aa;">${playerName} ←</td><td style="padding:6px 10px;text-align:right;font-weight:bold;color:#e8e8e8;">${myStanding.total}</td></tr>`;
  }

  return `<html>
    <head>
      <meta name="color-scheme" content="light dark">
      <meta name="supported-color-schemes" content="light dark">
    </head>
    <body style="background-color:#0d0d0d !important;margin:0;padding:0;">
      <div style="background-color:#0d0d0d;margin:0;padding:20px;font-family:'Barlow',-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;">
        <div style="max-width:600px;margin:0 auto;background:#161616;border-radius:8px;overflow:hidden;">
          <div style="background:linear-gradient(135deg, #1a472a 0%, #0a1f12 100%);color:white;padding:28px 20px;text-align:center;border-bottom:3px solid #00d4aa;">
            <h1 style="margin:0;font-size:28px;font-weight:800;text-transform:uppercase;letter-spacing:3px;">🏁 ${poolName}</h1>
            <p style="margin:10px 0 0;font-size:16px;opacity:0.85;letter-spacing:0.5px;">Round ${race.round}: ${race.name} — Results</p>
          </div>
          <div style="background:#161616;padding:32px 24px;color:#e8e8e8;">
            <p style="margin:0 0 16px;font-size:15px;">Hey <strong style="color:#00d4aa;">${playerName}</strong>,</p>
            <p style="margin:0 0 20px;font-size:15px;line-height:1.6;">The scores are in for the <strong>${race.name}</strong>!</p>
            ${raceWinner ? `<div style="background:rgba(0, 212, 170, 0.15);border-left:4px solid #00d4aa;border-radius:6px;padding:16px;margin:20px 0;text-align:center;"><strong style="font-size:17px;color:#00d4aa;">🏆 Race Winner: ${raceWinner.name} — ${raceWinner.total} pts ($25)</strong></div>` : ""}
            ${mySection}
            <div style="margin:24px 0;text-align:center;">
              <a href="${resultsSheetUrl}" style="display:inline-block;background:#00d4aa;color:#0d0d0d;text-decoration:none;font-size:14px;font-weight:800;padding:12px 20px;border-radius:6px;">View All Picks &amp; Full Results</a>
              <p style="margin:10px 0 0;color:#777;font-size:12px;line-height:1.5;">Open the race results sheet to see everyone&rsquo;s picks and scoring breakdown.</p>
            </div>
            <h3 style="color:#00d4aa;margin:24px 0 12px;font-size:15px;font-weight:700;text-transform:uppercase;letter-spacing:1px;">Race Scoreboard (Top 10)</h3>
            <table style="width:100%;border-collapse:collapse;font-size:13px;background:#0d0d0d;border:1px solid #2a2a2a;border-radius:6px;overflow:hidden;">
              <tr style="background:#1a472a;color:white;"><th style="padding:8px 12px;text-align:left;width:40px;">#</th><th style="padding:8px 12px;text-align:left;">Player</th><th style="padding:8px 12px;text-align:right;">Points</th></tr>
              ${top10Rows}
              ${allResults.length > 10 ? `<tr><td colspan="3" style="padding:8px 12px;color:#777;font-size:12px;background:#0d0d0d;">...and ${allResults.length - 10} more</td></tr>` : ""}
            </table>
            ${actualFinishers.length > 0 ? `<h3 style="color:#00d4aa;margin:24px 0 12px;font-size:15px;font-weight:700;text-transform:uppercase;letter-spacing:1px;">Actual Race Result</h3><p style="font-size:13px;color:#b8b8b8;line-height:1.8;background:#0d0d0d;padding:12px;border-radius:6px;border:1px solid #2a2a2a;">${actualPole ? `<strong style="color:#d4a017;">Pole:</strong> <span style="color:#e8e8e8;">${actualPole}</span> &nbsp;|&nbsp; ` : ""}${resultsList}</p>` : ""}
            <h3 style="color:#00d4aa;margin:24px 0 12px;font-size:15px;font-weight:700;text-transform:uppercase;letter-spacing:1px;">Season Standings</h3>
            <table style="width:100%;border-collapse:collapse;font-size:13px;background:#0d0d0d;border:1px solid #2a2a2a;border-radius:6px;overflow:hidden;">
              <tr style="background:#1a472a;color:white;"><th style="padding:6px 10px;text-align:left;width:40px;">#</th><th style="padding:6px 10px;text-align:left;">Player</th><th style="padding:6px 10px;text-align:right;">Total</th></tr>
              ${standingsRows}${myStandingRow}
            </table>
          </div>
          <div style="background:#0a0a0a;color:#777;padding:16px;text-align:center;font-size:12px;">2026 F1 Pool — See you at the next race! 🏎️</div>
        </div>
      </div>
    </body>
    </html>`;
}

// ── LEADERBOARD ─────────────────────────────────────────────────────────────

function recalculateLeaderboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lb = ss.getSheetByName("Leaderboard") || ss.insertSheet("Leaderboard");
  lb.clear();

  const scoredRaces = [];
  RACES_2026.forEach(race => {
    if (ss.getSheetByName(`R${race.round} Scores`)) scoredRaces.push(race);
  });

  if (scoredRaces.length === 0) { lb.getRange(1, 1).setValue("No races scored yet."); return; }

  const stats = {};

  scoredRaces.forEach(race => {
    const sheet = ss.getSheetByName(`R${race.round} Scores`);
    const data = sheet.getDataRange().getValues();
    let high = 0;

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0] === "ACTUAL RESULTS") break;
      const t = parseFloat(data[i][data[i].length - 1]) || 0;
      if (t > high) high = t;
    }

    for (let i = 1; i < data.length; i++) {
      let name = (data[i][0] || "").toString().trim();
      if (!name || name === "ACTUAL RESULTS") break;
      const clean = name.replace(/ \*$/, "");
      if (!stats[clean]) stats[clean] = { total: 0, wins: 0, best: 0, raceScores: {} };
      const t = parseFloat(data[i][data[i].length - 1]) || 0;
      stats[clean].total += t;
      stats[clean].raceScores[race.round] = t;
      if (t > stats[clean].best) stats[clean].best = t;
      if (t === high && high > 0) stats[clean].wins++;
    }
  });

  const sorted = Object.entries(stats).sort((a, b) => b[1].total - a[1].total);

  const baseHeaders = ["Rank", "Player", "Total Points", "Race Wins ($25)", "Best Race"];
  const raceHeaders = scoredRaces.map(r => `R${r.round}`);
  const allHeaders = [...baseHeaders, ...raceHeaders];
  const totalCols = allHeaders.length;

  lb.getRange(1, 1, 1, totalCols).setValues([allHeaders])
    .setFontWeight("bold").setBackground(HEADER_BG).setFontColor(HEADER_FG).setHorizontalAlignment("center");

  const colors = ["#ffd700", "#c0c0c0", "#cd7f32", "#d4edda", "#d4edda", "#d4edda"];
  sorted.forEach(([name, s], i) => {
    const row = i + 2;
    const rowData = [i + 1, name, Math.round(s.total * 10) / 10, s.wins, s.best];
    scoredRaces.forEach(race => rowData.push(s.raceScores[race.round] || 0));
    lb.getRange(row, 1, 1, totalCols).setValues([rowData]);
    if (i < 6) lb.getRange(row, 1, 1, totalCols).setBackground(colors[i]);
  });

  lb.setColumnWidth(1, 50); lb.setColumnWidth(2, 160); lb.setColumnWidth(3, 100);
  lb.setColumnWidth(4, 120); lb.setColumnWidth(5, 90);
  for (let c = 6; c <= totalCols; c++) lb.setColumnWidth(c, 50);

  if (sorted.length > 0) {
    lb.getRange(2, 1, sorted.length, 1).setHorizontalAlignment("center");
    lb.getRange(2, 3, sorted.length, totalCols - 2).setHorizontalAlignment("center");
  }

  lb.setFrozenRows(1);
  lb.setFrozenColumns(2);
}

// ── UTILITIES ───────────────────────────────────────────────────────────────

function getPlayers(ss) {
  const ps = ss.getSheetByName("Players");
  if (!ps || ps.getLastRow() < 2) return [];
  const data = ps.getDataRange().getValues();
  const players = [];
  for (let i = 1; i < data.length; i++) {
    const name = (data[i][0] || "").toString().trim();
    const email = (data[i][1] || "").toString().trim();
    if (name) players.push({ name, email });
  }
  return players;
}

function getDeadline(race) {
  const raceDate = new Date(race.date + "T00:00:00");
  const d = new Date(raceDate);
  d.setDate(d.getDate() - 3);
  d.setHours(20, 0, 0, 0);
  return d;
}

function findSheet(ss, prefix) {
  for (const s of ss.getSheets()) {
    if (s.getName().startsWith(prefix)) return s.getName();
  }
  return null;
}

function extractCode(val) {
  const s = (val || "").toString().trim().toUpperCase();
  if (s.length === 3 && DRIVERS_2026.includes(s)) return s;
  const match = s.match(/^([A-Z]{3})\s*-/);
  if (match && DRIVERS_2026.includes(match[1])) return match[1];
  const first3 = s.substring(0, 3);
  if (DRIVERS_2026.includes(first3)) return first3;
  return s;
}

function viewDeadlineStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cs = ss.getSheetByName("Race Calendar");
  if (!cs) { SpreadsheetApp.getUi().alert("Run Initial Setup first."); return; }
  const data = cs.getDataRange().getValues();
  const now = new Date();
  let status = "DEADLINE STATUS\n═══════════════════════════\n\n";
  for (let i = 1; i < data.length; i++) {
    const round = data[i][0], name = data[i][1];
    const deadline = new Date(data[i][4]);
    const hasPicks = data[i][5] === "Yes";
    const locked = data[i][7] === "Yes";
    const scored = data[i][8] === "Yes";
    let icon;
    if (scored) icon = "✅ Scored";
    else if (locked) icon = "🔒 Locked (ready to score)";
    else if (now > deadline) icon = "⏰ Past deadline" + (hasPicks ? " — LOCK NOW" : " (NO PICKS!)");
    else if (hasPicks) icon = "📧 Picks sent — " + Math.ceil((deadline - now) / 86400000) + " days left";
    else icon = "📋 Needs picks";
    status += `R${round}: ${name} — ${icon}\n`;
  }
  SpreadsheetApp.getUi().alert("Deadline Status", status, SpreadsheetApp.getUi().ButtonSet.OK);
}

function viewSubmissionStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Submission Status", "Enter round number:", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const round = parseInt(r.getResponseText());
  const race = RACES_2026.find(rc => rc.round === round);
  if (!race) { ui.alert("Invalid round."); return; }

  let respSheet = null;
  ss.getSheets().forEach(s => {
    if (s.getName().includes(`_R${round} Responses`)) respSheet = s;
  });

  const players = getPlayers(ss);
  const submitted = new Set();
  if (respSheet) {
    const data = respSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const name = (data[i][1] || "").toString().trim();
      if (name) submitted.add(name);
    }
  }

  let msg = `SUBMISSIONS: Round ${round} — ${race.name}\n═══════════════════════════\n\n`;
  let subCount = 0;
  players.forEach(p => {
    const did = submitted.has(p.name);
    if (did) subCount++;
    msg += `${did ? "✅" : "❌"} ${p.name}\n`;
  });
  msg += `\n${subCount} / ${players.length} submitted`;
  ui.alert("Submission Status", msg, ui.ButtonSet.OK);
}
