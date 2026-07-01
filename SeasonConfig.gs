// Edit this JSON before running initialSetup(). It is used only to seed the
// Settings, Teams, Drivers, and Race Calendar sheets; those sheets become the live
// source of truth after initialization.
const SEASON_CONFIG_JSON = `{
  "seasonYear": 2026,
  "poolName": "F1 Pool",
  "timeZone": "America/Vancouver",
  "deadlineDaysBefore": 3,
  "deadlineHour": 20,
  "teams": [
    { "name": "Red Bull", "color": "#3671C6" },
    { "name": "McLaren", "color": "#FF8000" },
    { "name": "Mercedes", "color": "#27F4D2" },
    { "name": "Ferrari", "color": "#E8002D" },
    { "name": "Williams", "color": "#64C4FF" },
    { "name": "Racing Bulls", "color": "#6692FF" },
    { "name": "Aston Martin", "color": "#229971" },
    { "name": "Haas", "color": "#B6BABD" },
    { "name": "Audi", "color": "#FF4C4C" },
    { "name": "Alpine", "color": "#FF87BC" },
    { "name": "Cadillac", "color": "#FFD700" }
  ],
  "drivers": [
    { "code": "VER", "name": "Verstappen", "team": "Red Bull" },
    { "code": "HAD", "name": "Hadjar", "team": "Red Bull" },
    { "code": "NOR", "name": "Norris", "team": "McLaren" },
    { "code": "PIA", "name": "Piastri", "team": "McLaren" },
    { "code": "RUS", "name": "Russell", "team": "Mercedes" },
    { "code": "ANT", "name": "Antonelli", "team": "Mercedes" },
    { "code": "LEC", "name": "Leclerc", "team": "Ferrari" },
    { "code": "HAM", "name": "Hamilton", "team": "Ferrari" },
    { "code": "SAI", "name": "Sainz", "team": "Williams" },
    { "code": "ALB", "name": "Albon", "team": "Williams" },
    { "code": "LAW", "name": "Lawson", "team": "Racing Bulls" },
    { "code": "LIN", "name": "Lindblad", "team": "Racing Bulls" },
    { "code": "ALO", "name": "Alonso", "team": "Aston Martin" },
    { "code": "STR", "name": "Stroll", "team": "Aston Martin" },
    { "code": "OCO", "name": "Ocon", "team": "Haas" },
    { "code": "BEA", "name": "Bearman", "team": "Haas" },
    { "code": "HUL", "name": "Hulkenberg", "team": "Audi" },
    { "code": "BOR", "name": "Bortoleto", "team": "Audi" },
    { "code": "GAS", "name": "Gasly", "team": "Alpine" },
    { "code": "COL", "name": "Colapinto", "team": "Alpine" },
    { "code": "PER", "name": "Perez", "team": "Cadillac" },
    { "code": "BOT", "name": "Bottas", "team": "Cadillac" }
  ],
  "races": [
    { "round": 1, "name": "Australian GP", "date": "2026-03-08", "day": "Sun" },
    { "round": 2, "name": "Chinese GP", "date": "2026-03-15", "day": "Sun" },
    { "round": 3, "name": "Japanese GP", "date": "2026-03-29", "day": "Sun" },
    { "round": 4, "name": "Miami GP", "date": "2026-05-03", "day": "Sun" },
    { "round": 5, "name": "Canadian GP", "date": "2026-05-24", "day": "Sun" },
    { "round": 6, "name": "Monaco GP", "date": "2026-06-07", "day": "Sun" },
    { "round": 7, "name": "Barcelona-Catalunya GP", "date": "2026-06-14", "day": "Sun" },
    { "round": 8, "name": "Austrian GP", "date": "2026-06-28", "day": "Sun" },
    { "round": 9, "name": "British GP", "date": "2026-07-05", "day": "Sun" },
    { "round": 10, "name": "Belgian GP", "date": "2026-07-19", "day": "Sun" },
    { "round": 11, "name": "Hungarian GP", "date": "2026-07-26", "day": "Sun" },
    { "round": 12, "name": "Dutch GP", "date": "2026-08-23", "day": "Sun" },
    { "round": 13, "name": "Italian GP", "date": "2026-09-06", "day": "Sun" },
    { "round": 14, "name": "Spanish GP", "date": "2026-09-13", "day": "Sun" },
    { "round": 15, "name": "Azerbaijan GP", "date": "2026-09-26", "day": "Sat" },
    { "round": 16, "name": "Singapore GP", "date": "2026-10-11", "day": "Sun" },
    { "round": 17, "name": "United States GP", "date": "2026-10-25", "day": "Sun" },
    { "round": 18, "name": "Mexico City GP", "date": "2026-11-01", "day": "Sun" },
    { "round": 19, "name": "Brazilian GP", "date": "2026-11-08", "day": "Sun" },
    { "round": 20, "name": "Las Vegas GP", "date": "2026-11-21", "day": "Sat" },
    { "round": 21, "name": "Qatar GP", "date": "2026-11-29", "day": "Sun" },
    { "round": 22, "name": "Abu Dhabi GP", "date": "2026-12-06", "day": "Sun" }
  ]
}`;

function getInitialSeasonConfig() {
  let config;
  try {
    config = JSON.parse(SEASON_CONFIG_JSON);
  } catch (e) {
    throw new Error(`SeasonConfig.gs contains invalid JSON: ${e.message}`);
  }

  if (!Number.isInteger(config.seasonYear)) throw new Error("seasonYear must be an integer.");
  if (!config.poolName || !config.timeZone) throw new Error("poolName and timeZone are required.");
  if (!Number.isInteger(config.deadlineDaysBefore) || !Number.isInteger(config.deadlineHour)) {
    throw new Error("deadlineDaysBefore and deadlineHour must be integers.");
  }
  if (!Array.isArray(config.teams) || config.teams.length === 0) throw new Error("At least one team is required.");
  if (!Array.isArray(config.drivers) || config.drivers.length < 10) throw new Error("At least 10 drivers are required.");
  if (!Array.isArray(config.races) || config.races.length === 0) throw new Error("At least one race is required.");

  const teamNames = new Set();
  config.teams.forEach(team => {
    team.name = String(team.name || "").trim();
    team.color = String(team.color || "").trim().toUpperCase();
    if (!team.name || !/^#[0-9A-F]{6}$/.test(team.color)) throw new Error("Every team needs a unique name and a 6-digit hex color.");
    if (teamNames.has(team.name)) throw new Error(`Duplicate team name: ${team.name}`);
    teamNames.add(team.name);
  });

  const driverCodes = new Set();
  config.drivers.forEach(driver => {
    driver.code = String(driver.code || "").trim().toUpperCase();
    if (!/^[A-Z]{3}$/.test(driver.code) || !driver.name || !driver.team) throw new Error("Every driver needs a unique 3-letter code, name, and team.");
    if (driverCodes.has(driver.code)) throw new Error(`Duplicate driver code: ${driver.code}`);
    if (!teamNames.has(driver.team)) throw new Error(`Unknown team for ${driver.code}: ${driver.team}`);
    driverCodes.add(driver.code);
  });

  const rounds = new Set();
  config.races.forEach(race => {
    if (!Number.isInteger(race.round) || !race.name || !/^\d{4}-\d{2}-\d{2}$/.test(race.date)) {
      throw new Error("Every race needs an integer round, name, and YYYY-MM-DD date.");
    }
    if (rounds.has(race.round)) throw new Error(`Duplicate race round: ${race.round}`);
    rounds.add(race.round);
  });

  return config;
}
