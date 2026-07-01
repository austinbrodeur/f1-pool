// =============================================================================
// PUBLIC RESULTS WORKBOOK
// Creates and refreshes a separate viewer-safe spreadsheet containing public
// race result tabs, a season summary, and the two charts from the legacy pool.
// =============================================================================

const PUBLIC_RESULTS_ID_PROPERTY = "public_results_spreadsheet_id";
const PUBLIC_RESULTS_LINKS_PROPERTY = "public_results_links";
const PUBLIC_RESULTS_SYNCED_AT_PROPERTY = "public_results_synced_at";

function rebuildPublicResults() {
  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  try {
    const links = syncPublicResults_(sourceSs);
    ui.alert(
      "Public Results Updated",
      `The viewer workbook is ready:\n\n${links.summaryUrl}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert("Public Results Error", e.message, ui.ButtonSet.OK);
    throw e;
  }
}

function syncPublicResults_(sourceSs) {
  const settings = getPoolSettings(sourceSs);
  const races = getRaces(sourceSs);
  const publicSs = getOrCreatePublicResultsSpreadsheet_(settings);
  syncPublicViewers_(publicSs, getPlayers(sourceSs));

  const raceResults = [];
  const playerNames = new Set(getPlayers(sourceSs).map(player => player.name));

  races.forEach(race => {
    const sourceSheet = sourceSs.getSheetByName(`R${race.round} Scores`);
    if (!sourceSheet) return;
    const result = readRaceScoreData_(sourceSheet, race);
    result.players.forEach(player => playerNames.add(player.name));
    raceResults.push(result);
  });

  const model = buildPublicResultsModel_(Array.from(playerNames), races, raceResults);
  const summarySheet = ensurePublicSheet_(publicSs, "Summary");
  const ourRaceSheet = ensurePublicSheet_(publicSs, "Our Race");
  const pointsSheet = ensurePublicSheet_(publicSs, "Points Data Graph");

  writePublicSummary_(summarySheet, settings, model);
  writeOurRaceChart_(ourRaceSheet, model);
  writePointsChart_(pointsSheet, model);

  const raceUrls = {};
  raceResults.forEach(result => {
    const sheetName = makePublicRaceSheetName_(result.race);
    const targetSheet = ensurePublicSheet_(publicSs, sheetName);
    writePublicRaceSheet_(targetSheet, result);
    raceUrls[result.race.round] = `${publicSs.getUrl()}#gid=${targetSheet.getSheetId()}`;
  });

  const pickUrls = {};
  races.forEach(race => {
    const sourcePickSheetName = findSheet(sourceSs, `R${race.round} Picks`);
    if (!sourcePickSheetName) return;
    const sourcePickSheet = sourceSs.getSheetByName(sourcePickSheetName);
    const targetSheet = ensurePublicSheet_(publicSs, makePublicPickSheetName_(race));
    writePublicPickSheet_(targetSheet, sourcePickSheet.getDataRange().getValues());
    pickUrls[race.round] = `${publicSs.getUrl()}#gid=${targetSheet.getSheetId()}`;
  });

  movePublicSheet_(publicSs, summarySheet, 1);
  movePublicSheet_(publicSs, ourRaceSheet, 2);
  movePublicSheet_(publicSs, pointsSheet, 3);
  SpreadsheetApp.flush();

  const links = {
    spreadsheetId: publicSs.getId(),
    summaryUrl: `${publicSs.getUrl()}#gid=${summarySheet.getSheetId()}`,
    raceUrls: raceUrls,
    pickUrls: pickUrls
  };
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PUBLIC_RESULTS_LINKS_PROPERTY, JSON.stringify(links));
  props.setProperty(PUBLIC_RESULTS_SYNCED_AT_PROPERTY, String(Date.now()));
  return links;
}

function getFreshPublicResultsLinks_(sourceSs, maxAgeMilliseconds) {
  const props = PropertiesService.getScriptProperties();
  const syncedAt = parseInt(props.getProperty(PUBLIC_RESULTS_SYNCED_AT_PROPERTY), 10);
  const linksJson = props.getProperty(PUBLIC_RESULTS_LINKS_PROPERTY);
  if (linksJson && Number.isFinite(syncedAt) && Date.now() - syncedAt <= maxAgeMilliseconds) {
    try { return JSON.parse(linksJson); } catch (e) {}
  }
  return syncPublicResults_(sourceSs);
}

function getOrCreatePublicResultsSpreadsheet_(settings) {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(PUBLIC_RESULTS_ID_PROPERTY);
  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (e) {
      Logger.log(`Stored public results workbook is unavailable: ${e.message}`);
      props.deleteProperty(PUBLIC_RESULTS_ID_PROPERTY);
    }
  }

  const ss = SpreadsheetApp.create(`${settings.poolName} ${settings.seasonYear} — Public Results`);
  props.setProperty(PUBLIC_RESULTS_ID_PROPERTY, ss.getId());
  ss.setSpreadsheetTimeZone(settings.timeZone);
  ss.getSheets()[0].setName("Summary");

  try {
    DriveApp.getFileById(ss.getId()).setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
  } catch (e) {
    Logger.log(`Anyone-with-link sharing was not available: ${e.message}`);
  }

  return ss;
}

function syncPublicViewers_(publicSs, players) {
  const emails = [...new Set(players.map(player => player.email).filter(Boolean))];
  if (emails.length === 0) return;
  try {
    publicSs.addViewers(emails);
  } catch (e) {
    Logger.log(`Bulk viewer sync failed; retrying individually: ${e.message}`);
    emails.forEach(email => {
      try { publicSs.addViewer(email); }
      catch (viewerError) { Logger.log(`Could not add ${email} as a viewer: ${viewerError.message}`); }
    });
  }
}

function ensurePublicSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function movePublicSheet_(ss, sheet, position) {
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(position);
}

function makePublicRaceSheetName_(race) {
  return `R${race.round} Results — ${race.name}`.slice(0, 99);
}

function makePublicPickSheetName_(race) {
  return `R${race.round} Picks — ${race.name}`.slice(0, 99);
}

function readRaceScoreData_(sheet, race) {
  const values = sheet.getDataRange().getValues();
  const players = [];
  const totalColumn = values.length > 0 ? values[0].length - 1 : 0;
  for (let i = 1; i < values.length; i++) {
    const displayName = String(values[i][0] || "").trim();
    if (!displayName || displayName === "ACTUAL RESULTS") break;
    players.push({
      name: displayName.replace(/ \*$/, ""),
      displayName: displayName,
      score: parseFloat(values[i][totalColumn]) || 0
    });
  }
  return { race: race, values: values, players: players };
}

function buildPublicResultsModel_(playerNames, races, raceResults) {
  const names = [...new Set(playerNames.filter(Boolean))];
  const resultByRound = {};
  raceResults.forEach(result => { resultByRound[result.race.round] = result; });

  const scoreByRound = {};
  raceResults.forEach(result => {
    const scores = {};
    result.players.forEach(player => { scores[player.name] = player.score; });
    scoreByRound[result.race.round] = scores;
  });

  const totals = {};
  names.forEach(name => { totals[name] = 0; });
  raceResults.forEach(result => {
    result.players.forEach(player => { totals[player.name] = (totals[player.name] || 0) + player.score; });
  });

  const sortedPlayers = names.slice().sort((a, b) => (totals[b] - totals[a]) || a.localeCompare(b));
  const cumulative = {};
  const ranksByRound = {};
  names.forEach(name => { cumulative[name] = 0; });

  raceResults.forEach(result => {
    names.forEach(name => { cumulative[name] += (scoreByRound[result.race.round][name] || 0); });
    const ranked = names.slice().sort((a, b) => (cumulative[b] - cumulative[a]) || a.localeCompare(b));
    const ranks = {};
    ranked.forEach((name, index) => { ranks[name] = index + 1; });
    ranksByRound[result.race.round] = ranks;
  });

  const latestResult = raceResults.length > 0 ? raceResults[raceResults.length - 1] : null;
  return {
    races: races,
    raceResults: raceResults,
    resultByRound: resultByRound,
    scoreByRound: scoreByRound,
    totals: totals,
    sortedPlayers: sortedPlayers,
    ranksByRound: ranksByRound,
    latestTopScores: latestResult ? groupTopRaceScores_(latestResult.players, 5) : [],
    latestRace: latestResult ? latestResult.race : null,
    seasonHighScores: groupSeasonHighScores_(raceResults, 5)
  };
}

function groupTopRaceScores_(players, limit) {
  const grouped = {};
  players.forEach(player => {
    const key = String(player.score);
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(player.name);
  });
  return Object.keys(grouped)
    .map(Number)
    .sort((a, b) => b - a)
    .slice(0, limit)
    .map(score => ({ score: score, label: grouped[String(score)].sort().join(" & ") }));
}

function groupSeasonHighScores_(raceResults, limit) {
  const grouped = {};
  raceResults.forEach(result => {
    result.players.forEach(player => {
      const key = String(player.score);
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(`${player.name} (${result.race.name})`);
    });
  });
  return Object.keys(grouped)
    .map(Number)
    .sort((a, b) => b - a)
    .slice(0, limit)
    .map(score => ({ score: score, label: grouped[String(score)].sort().join(", ") }));
}

function resetPublicSheet_(sheet) {
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  sheet.getDataRange().breakApart();
  sheet.clear();
  sheet.setHiddenGridlines(true);
}

function ensurePublicGridSize_(sheet, rows, columns) {
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < columns) sheet.insertColumnsAfter(sheet.getMaxColumns(), columns - sheet.getMaxColumns());
}

function writePublicSummary_(sheet, settings, model) {
  resetPublicSheet_(sheet);
  const scoredRaces = model.raceResults.map(result => result.race);
  const totalColumn = model.races.length + 3;
  const positionStartColumn = totalColumn + 2;
  const tableRows = Math.max(model.sortedPlayers.length, 1);
  const calloutRow = tableRows + 7;
  const positionColumnCount = Math.max(scoredRaces.length + 1, 1);
  const usedRows = calloutRow + 5;
  const usedColumns = Math.max(positionStartColumn + positionColumnCount - 1, totalColumn, 11);
  ensurePublicGridSize_(sheet, usedRows, usedColumns);

  const title = `${settings.seasonYear} Formula 1 Race Pool Summary`;
  sheet.getRange(1, 1, 1, Math.max(totalColumn, 8))
    .setBackground("#1a472a").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(1, 2).setValue(title).setFontSize(18);

  const scoreHeaders = ["Rank", "Player", ...model.races.map(race => race.name), "TOTAL"];
  sheet.getRange(3, 1, 1, scoreHeaders.length).setValues([scoreHeaders]);
  const scoreRows = model.sortedPlayers.map((name, index) => [
    index + 1,
    name,
    ...model.races.map(race => (model.scoreByRound[race.round] || {})[name] || 0),
    Math.round((model.totals[name] || 0) * 10) / 10
  ]);
  if (scoreRows.length > 0) sheet.getRange(4, 1, scoreRows.length, scoreHeaders.length).setValues(scoreRows);

  const positionHeaders = ["Player", ...scoredRaces.map(race => race.name)];
  sheet.getRange(1, positionStartColumn, 1, positionHeaders.length).merge().setValue("Position After Race to Date")
    .setBackground("#1a472a").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.getRange(3, positionStartColumn, 1, positionHeaders.length).setValues([positionHeaders]);
  const positionRows = model.sortedPlayers.map(name => [
    name,
    ...scoredRaces.map(race => model.ranksByRound[race.round][name])
  ]);
  if (positionRows.length > 0) sheet.getRange(4, positionStartColumn, positionRows.length, positionHeaders.length).setValues(positionRows);

  const latestTitle = model.latestRace ? `Top 5 for ${model.latestRace.name}` : "Top 5 for Latest Race";
  sheet.getRange(calloutRow, 1).setValue("Pos.");
  sheet.getRange(calloutRow, 2, 1, 3).merge().setValue(latestTitle);
  sheet.getRange(calloutRow, 5).setValue("Pts");
  Array.from({ length: 5 }, (_, index) => {
    const item = model.latestTopScores[index];
    return [index + 1, item ? item.label : "", item ? item.score : ""];
  }).forEach((row, index) => {
    const targetRow = calloutRow + index + 1;
    sheet.getRange(targetRow, 1).setValue(row[0]);
    sheet.getRange(targetRow, 2, 1, 3).merge().setValue(row[1]);
    sheet.getRange(targetRow, 5).setValue(row[2]);
  });

  sheet.getRange(calloutRow, 7).setValue("Pos.");
  sheet.getRange(calloutRow, 8, 1, 3).merge().setValue("Season Race High Scores");
  sheet.getRange(calloutRow, 11).setValue("Pts");
  Array.from({ length: 5 }, (_, index) => {
    const item = model.seasonHighScores[index];
    return [index + 1, item ? item.label : "", item ? item.score : ""];
  }).forEach((row, index) => {
    const targetRow = calloutRow + index + 1;
    sheet.getRange(targetRow, 7).setValue(row[0]);
    sheet.getRange(targetRow, 8, 1, 3).merge().setValue(row[1]);
    sheet.getRange(targetRow, 11).setValue(row[2]);
  });

  sheet.getRange(2, 1, usedRows - 1, usedColumns).setBackground("#161616").setFontColor("#e8e8e8");
  sheet.getRange(3, 1, 1, scoreHeaders.length).setBackground("#1a472a").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(3, positionStartColumn, 1, positionHeaders.length).setBackground("#1a472a").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(calloutRow, 1, 1, 5).setBackground("#1a472a").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(calloutRow, 7, 1, 5).setBackground("#1a472a").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(calloutRow + 1, 2, 5, 3).setWrap(true);
  sheet.getRange(calloutRow + 1, 8, 5, 3).setWrap(true);
  if (scoreRows.length > 0) {
    sheet.getRange(4, 1, scoreRows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(4, 3, scoreRows.length, scoreHeaders.length - 2).setHorizontalAlignment("center");
    sheet.getRange(4, 1, 1, scoreHeaders.length).setBackground("#d4a017").setFontColor("#0d0d0d").setFontWeight("bold");
  }
  sheet.setFrozenRows(3);
  // Freeze before column B so no merged callout range crosses the boundary.
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 170);
  for (let col = 3; col <= scoreHeaders.length; col++) sheet.setColumnWidth(col, 85);
  sheet.setColumnWidth(positionStartColumn, 170);
  for (let col = positionStartColumn + 1; col < positionStartColumn + positionHeaders.length; col++) sheet.setColumnWidth(col, 85);
  sheet.setColumnWidth(2, 190);
}

function writeOurRaceChart_(sheet, model) {
  resetPublicSheet_(sheet);
  if (model.raceResults.length === 0 || model.sortedPlayers.length === 0) {
    sheet.getRange("A1").setValue("Position history will appear after the first scored race.");
    return;
  }

  const dataColumn = 27;
  const header = ["Race", ...model.sortedPlayers];
  const rows = model.raceResults.map(result => [
    result.race.name,
    ...model.sortedPlayers.map(name => model.ranksByRound[result.race.round][name])
  ]);
  ensurePublicGridSize_(sheet, rows.length + 1, dataColumn + header.length - 1);
  sheet.getRange(1, dataColumn, 1, header.length).setValues([header]);
  sheet.getRange(2, dataColumn, rows.length, header.length).setValues(rows);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, dataColumn, rows.length + 1, header.length))
    .setNumHeaders(1)
    .setPosition(1, 1, 0, 0)
    .setOption("title", "Position After Race to Date")
    .setOption("width", 1250)
    .setOption("height", 720)
    .setOption("legend", { position: "right", textStyle: { fontSize: 9 } })
    .setOption("pointSize", 3)
    .setOption("vAxis", { direction: -1, viewWindow: { min: 1, max: model.sortedPlayers.length } })
    .setOption("hAxis", { slantedText: true, slantedTextAngle: 35 })
    .build();
  sheet.insertChart(chart);
}

function writePointsChart_(sheet, model) {
  resetPublicSheet_(sheet);
  if (model.sortedPlayers.length === 0) {
    sheet.getRange("A1").setValue("Point standings will appear after players are added.");
    return;
  }

  const dataColumn = 27;
  const rows = [["Player", "Points"], ...model.sortedPlayers.map(name => [name, model.totals[name] || 0])];
  ensurePublicGridSize_(sheet, rows.length, dataColumn + 1);
  sheet.getRange(1, dataColumn, rows.length, 2).setValues(rows);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(1, dataColumn, rows.length, 2))
    .setNumHeaders(1)
    .setPosition(1, 1, 0, 0)
    .setOption("title", "Player Points Standings")
    .setOption("width", 1050)
    .setOption("height", Math.max(650, model.sortedPlayers.length * 24))
    .setOption("legend", { position: "none" })
    .setOption("colors", ["#00d4aa"])
    .setOption("hAxis", { title: "Points", minValue: 0 })
    .setOption("vAxis", { direction: -1 })
    .build();
  sheet.insertChart(chart);
}

function writePublicRaceSheet_(sheet, result) {
  resetPublicSheet_(sheet);
  if (result.values.length === 0 || result.values[0].length === 0) return;
  const rows = result.values.length;
  const columns = result.values[0].length;
  ensurePublicGridSize_(sheet, rows, columns);
  sheet.getRange(1, 1, rows, columns).setValues(result.values);
  sheet.getRange(1, 1, 1, columns).setBackground("#1a472a").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(2, 1, Math.max(rows - 1, 1), columns).setBackground("#161616").setFontColor("#e8e8e8");
  if (result.players.length > 0) {
    sheet.getRange(2, 1, 1, columns).setBackground("#d4a017").setFontColor("#0d0d0d").setFontWeight("bold");
  }
  for (let i = 1; i < result.values.length; i++) {
    if (String(result.values[i][0] || "") === "ACTUAL RESULTS") {
      sheet.getRange(i + 1, 1, 1, Math.min(columns, 2)).setBackground("#1a472a").setFontColor("#ffffff").setFontWeight("bold");
      break;
    }
  }
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.autoResizeColumns(1, columns);
  sheet.setColumnWidth(1, 180);
}

function writePublicPickSheet_(sheet, values) {
  resetPublicSheet_(sheet);
  if (values.length === 0 || values[0].length === 0) return;
  const rows = values.length;
  const columns = values[0].length;
  ensurePublicGridSize_(sheet, rows, columns);
  sheet.getRange(1, 1, rows, columns).setValues(values);
  sheet.getRange(1, 1, 1, columns).setBackground("#0d0d0d").setFontColor("#ffffff").setFontWeight("bold");
  if (rows > 1) {
    sheet.getRange(2, 1, 1, columns).setBackground("#1a472a").setFontColor("#ffffff").setFontWeight("bold");
  }
  if (rows > 2) {
    sheet.getRange(3, 1, rows - 2, columns).setBackground("#161616").setFontColor("#e8e8e8");
  }
  sheet.setFrozenRows(Math.min(2, rows));
  sheet.setFrozenColumns(1);
  sheet.autoResizeColumns(1, columns);
  sheet.setColumnWidth(1, 180);
  if (columns >= 2) sheet.setColumnWidth(2, 160);
}
