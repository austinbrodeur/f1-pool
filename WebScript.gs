// =============================================================================
// WEB APP PICK PAGE — Add this to your Apps Script project as a new file
// (In Apps Script: File > New > Script file, name it "WebApp")
//
// This replaces Google Forms with a custom pick page served as an Apps Script Web App.
// Players get a unique URL with a token, see a polished pick interface with
// real-time duplicate prevention, and submit directly to the spreadsheet.
//
// SETUP:
//   1. Add this file to your Apps Script project
//   2. Add the PickPage.html file (create via File > New > HTML file, name it "PickPage")
//   3. Deploy: Deploy > New deployment > Web app > Anyone can access > Deploy
//   4. Copy the web app URL — this is your base URL
//   5. Update WEB_APP_URL below with your deployed URL
//   6. Use "Create & Email Web Picks" from the F1 Pool menu
// =============================================================================

// ⚠️ UPDATE THIS after deploying the web app
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxNFKbYMlCdC30xh1Rm8DHB9rcBetWfiDS8JzBEuu_e1Gf6cBFc4K6BnmDXqStHR16P_w/exec"; // e.g. "https://script.google.com/macros/s/xxxxx/exec"

// ── WEB APP ENTRY POINT ─────────────────────────────────────────────────────

function doGet(e) {
  const token = (e && e.parameter && e.parameter.t) || "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let poolTitle = "F1 Pool";
  try {
    const settings = getPoolSettings(ss);
    poolTitle = `${settings.poolName} ${settings.seasonYear}`;
  } catch (err) {}

  // If no token, show a friendly message instead of an error
  if (!token) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;max-width:500px;margin:60px auto;text-align:center;">' +
      `<h2 style="color:#1a472a;">${poolTitle}</h2>` +
      '<p>This page requires a personal pick link.</p>' +
      '<p style="color:#666;">Check your email for your unique link to submit picks.</p>' +
      '</div>'
    ).setTitle(poolTitle);
  }

  const template = HtmlService.createTemplateFromFile("PickPage");
  template.token = token;
  return template.evaluate()
    .setTitle(`${poolTitle} — Submit Picks`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ── TOKEN MANAGEMENT ────────────────────────────────────────────────────────

// Stable secret stored in script properties (created once, reused forever)
function getTokenSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty("token_secret");
  if (!secret) {
    secret = Utilities.getUuid();
    props.setProperty("token_secret", secret);
  }
  return secret;
}

function generateToken(playerName, roundNum) {
  const secret = getTokenSecret_();
  const raw = playerName + "|" + roundNum + "|" + secret;
  return Utilities.base64EncodeWebSafe(raw);
}

function decodeToken(token) {
  if (!token) { Logger.log("decodeToken: empty token"); return null; }

  Logger.log("decodeToken raw token: " + token);

  // Decode directly
  const bytes = Utilities.base64DecodeWebSafe(token);
  const decoded = Utilities.newBlob(bytes).getDataAsString();
  Logger.log("decodeToken decoded: " + decoded);

  // Parse: "playerName|roundNum|secret"
  const firstPipe = decoded.indexOf("|");
  const secondPipe = decoded.indexOf("|", firstPipe + 1);
  if (firstPipe === -1 || secondPipe === -1) { Logger.log("decodeToken: missing pipes"); return null; }

  const playerName = decoded.substring(0, firstPipe);
  const roundNum = parseInt(decoded.substring(firstPipe + 1, secondPipe));
  const tokenSecret = decoded.substring(secondPipe + 1);

  // Verify secret
  const expectedSecret = getTokenSecret_();
  if (tokenSecret !== expectedSecret) {
    Logger.log("decodeToken: secret mismatch");
    return null;
  }

  return { playerName, roundNum };
}

// ── SERVER-SIDE FUNCTIONS (called from the HTML page) ───────────────────────

function getPickPageData(token) {
  const info = decodeToken(token);
  if (!info) return { error: "Invalid or expired link." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getPoolSettings(ss);
  const race = getRaceByRound(ss, info.roundNum);
  if (!race) return { error: "Race not found." };

  const deadline = getDeadline(race, settings);
  const now = new Date();
  const isPastDeadline = now >= deadline;

  // Check for existing submission
  const props = PropertiesService.getScriptProperties();
  const picksJson = props.getProperty(`webpicks_R${info.roundNum}_${info.playerName}`);
  let existingPicks = null;
  if (picksJson) {
    try { existingPicks = JSON.parse(picksJson); } catch(e) {}
  }

  return {
    playerName: info.playerName,
    roundNum: info.roundNum,
    raceName: race.name,
    raceDate: race.date,
    poolTitle: `${settings.poolName} ${settings.seasonYear}`,
    deadline: deadline.toISOString(),
    deadlineDisplay: Utilities.formatDate(deadline, settings.timeZone, "EEEE, MMM d 'at' h:mm a"),
    isPastDeadline: isPastDeadline,
    drivers: getDrivers(ss),
    existingPicks: existingPicks
  };
}

function submitPicks(token, picks) {
  const info = decodeToken(token);
  if (!info) return { error: "Invalid or expired link." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getPoolSettings(ss);
  const driverCodes = getDrivers(ss).map(driver => driver.code);
  const race = getRaceByRound(ss, info.roundNum);
  if (!race) return { error: "Race not found." };

  // Check deadline
  const deadline = getDeadline(race, settings);
  if (new Date() >= deadline) {
    return { error: "The deadline has passed. Picks can no longer be submitted." };
  }

  // Validate picks
  if (!picks.pole || !picks.positions || picks.positions.length !== 10) {
    return { error: "Incomplete picks. Select a pole prediction and all 10 positions." };
  }

  // Check for duplicates in P1-P10
  const unique = [...new Set(picks.positions)];
  if (unique.length !== 10) {
    return { error: "Duplicate drivers in P1-P10. Each driver can only appear once." };
  }

  // Validate all codes exist
  const allCodes = [picks.pole, ...picks.positions];
  for (const code of allCodes) {
    if (!driverCodes.includes(code)) {
      return { error: `Unknown driver code: ${code}` };
    }
  }

  // Store picks in script properties (will be read by Lock & Import)
  const props = PropertiesService.getScriptProperties();
  const pickData = {
    playerName: info.playerName,
    pole: picks.pole,
    positions: picks.positions,
    timestamp: new Date().toISOString()
  };
  props.setProperty(`webpicks_R${info.roundNum}_${info.playerName}`, JSON.stringify(pickData));

  // Also write to the hidden response sheet for compatibility with existing lock/import
  writePickToResponseSheet(info.roundNum, info.playerName, picks);

  return { success: true, message: "Picks submitted! You can resubmit anytime before the deadline." };
}

function writePickToResponseSheet(roundNum, playerName, picks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Find or create hidden response sheet
  let respSheetName = `_R${roundNum} Responses (hidden)`;
  let respSheet = ss.getSheetByName(respSheetName);
  if (!respSheet) {
    respSheet = ss.insertSheet(respSheetName);
    respSheet.getRange(1, 1, 1, 13).setValues([[
      "Timestamp", "Player", "Pole",
      "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10"
    ]]);
    respSheet.hideSheet();
  }

  // Append the response (latest per player will be used at lock time)
  const row = [
    new Date(),
    playerName,
    picks.pole,
    ...picks.positions
  ];
  respSheet.appendRow(row);
}

// ── CREATE & EMAIL WEB PICK LINKS ───────────────────────────────────────────

function createWebPicksForNextRace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cs = ss.getSheetByName("Race Calendar");
  if (!cs) { SpreadsheetApp.getUi().alert("Run Initial Setup first."); return; }
  const data = cs.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === "No") { createAndEmailWebPicks(data[i][0]); return; }
  }
  SpreadsheetApp.getUi().alert("All races already have pick forms!");
}

function createWebPicksPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Create Web Picks", "Enter round number:", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  createAndEmailWebPicks(parseInt(r.getResponseText()));
}

function createAndEmailWebPicks(roundNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settings = getPoolSettings(ss);
  const race = getRaceByRound(ss, roundNum);
  if (!race) { ui.alert("Invalid round."); return; }

  if (WEB_APP_URL === "YOUR_WEB_APP_URL_HERE") {
    ui.alert("⚠️ Web App URL not set!\n\nDeploy the web app first:\n1. Deploy > New deployment > Web app\n2. Set 'Anyone' can access\n3. Copy the URL and update WEB_APP_URL in the WebApp.gs file.");
    return;
  }

  const players = getPlayers(ss);
  if (players.length === 0) { ui.alert("Add players to the Players sheet first!"); return; }
  const playersWithEmail = players.filter(p => p.email);
  if (playersWithEmail.length === 0) { ui.alert("No players have email addresses!"); return; }

  const deadline = getDeadline(race, settings);
  const deadlineStr = Utilities.formatDate(deadline, settings.timeZone, "EEEE, MMM d 'at' h:mm a");
  const poolName = `${settings.poolName} ${settings.seasonYear}`;
  const props = PropertiesService.getScriptProperties();

  // Create hidden response sheet
  let respSheetName = `_R${roundNum} Responses (hidden)`;
  let respSheet = ss.getSheetByName(respSheetName);
  if (!respSheet) {
    respSheet = ss.insertSheet(respSheetName);
    respSheet.getRange(1, 1, 1, 13).setValues([[
      "Timestamp", "Player", "Pole",
      "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10"
    ]]);
    respSheet.hideSheet();
  }

  let emailsSent = 0;
  const playerLinks = {};

  playersWithEmail.forEach(player => {
    try {
      const token = generateToken(player.name, roundNum);
      const pickUrl = WEB_APP_URL + "?t=" + token;
      playerLinks[player.name] = pickUrl;

      MailApp.sendEmail({
        to: player.email,
        subject: `${poolName} Picks — Round ${race.round}: ${race.name}`,
        htmlBody: buildPickEmail(player.name, race, deadlineStr, pickUrl, poolName)
      });
      emailsSent++;
    } catch(e) {
      Logger.log(`Failed to email ${player.name}: ${e.message}`);
    }
  });

  props.setProperty(`form_links_${roundNum}`, JSON.stringify(playerLinks));

  // Update calendar
  const cs = ss.getSheetByName("Race Calendar");
  const calendarRow = getCalendarRow(ss, roundNum);
  if (calendarRow > 0) {
    cs.getRange(calendarRow, 6).setValue("Yes");
    cs.getRange(calendarRow, 7).setValue("Web App");
  }

  ui.alert(
    `Web Picks Created & Emailed!`,
    `Round ${race.round}: ${race.name}\n\n` +
    `📧 Emails sent: ${emailsSent} / ${playersWithEmail.length}\n` +
    `⏰ Deadline: ${deadlineStr}\n\n` +
    `Each player received a UNIQUE link to the web pick page.\n` +
    `Picks are written directly to the response sheet.\n` +
    `After the deadline, use "Lock & Import" to reveal picks.`,
    ui.ButtonSet.OK
  );
}
